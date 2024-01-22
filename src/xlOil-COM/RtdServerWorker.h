#pragma once
#include "RtdManager.h"
#include <xloil/Log.h>

#include <atlbase.h>
#include <string>
#include <unordered_map>
#include <unordered_set>
#include <mutex>
#include <shared_mutex>
#include <memory>
#include <vector>

using std::vector;
using std::shared_ptr;
using std::scoped_lock;
using std::unique_lock;
using std::shared_lock;
using std::wstring;
using std::unordered_set;
using std::unordered_map;
using std::pair;
using std::list;
using std::atomic;
using std::mutex;

namespace xloil
{
  namespace COM
  {
    template <class TValue>
    class RtdServerThreadedWorker : public IRtdServerWorker, public IRtdPublishManager<TValue>
    {
    public:

      void start(std::function<void()>&& updateNotify)
      {
        _updateNotify = std::move(updateNotify);
        _isRunning = true;
        _workerThread = std::thread([=]() { this->workerThreadMain(); });
      }
      void connect(long topicId, wstring&& topic)
      {
        {
          unique_lock lock(_mutexNewSubscribers);
          _topicsToConnect.emplace_back(topicId, std::move(topic));
        }
        notify();
      }
      void disconnect(long topicId)
      {
        {
          unique_lock lock(_mutexNewSubscribers);
          _topicIdsToDisconnect.emplace_back(topicId);
        }
        notify();
      }

      SAFEARRAY* getUpdates() 
      { 
        auto updates = _readyUpdates.exchange(nullptr);
        notify();
        return updates;
      }

      void quit()
      {
        if (!isServerRunning())
          return; // Already terminated, or never started

        setQuitFlag();
        // Let thread know we have set 'quit' flag
        notify();
      }

      void join()
      {
        quit();
        if (_workerThread.joinable())
          _workerThread.join();
      }

      void update(wstring&& topic, const shared_ptr<TValue>& value)
      {
        if (!isServerRunning())
          return;
        {
          scoped_lock lock(_mutexNewValues);
          // TODO: can this be somehow lock free?
          _newValues.emplace_back(make_pair(std::move(topic), value));
        }
        notify();
      }

      void addPublisher(const shared_ptr<IRtdPublisher>& job)
      {
        auto existingJob = job;
        {
          unique_lock lock(_lockRecords);
          auto& record = _records[job->topic()];
          std::swap(record.publisher, existingJob);
          if (existingJob)
            _cancelledPublishers.push_back(existingJob);
        }
        if (existingJob)
          existingJob->stop();
      }

      bool dropPublisher(const wchar_t* topic)
      {
        // We must not hold the lock when calling functions on the publisher
        // as they may try to call other functions on the RTD server. 
        shared_ptr<IRtdPublisher> publisher;
        {
          unique_lock lock(_lockRecords);
          auto i = _records.find(topic);
          if (i == _records.end())
            return false;
          std::swap(publisher, i->second.publisher);
        }

        // Signal the publisher to stop
        publisher->stop();

        // Destroy producer, the dtor of RtdPublisher waits for completion
        publisher.reset();

        // Publish empty value (which triggers a notify)
        update(topic, shared_ptr<TValue>());
        return true;
      }

      bool value(const wchar_t* topic, shared_ptr<const TValue>& val) const
      {
        shared_lock lock(_lockRecords);
        auto found = _records.find(topic);
        if (found == _records.end())
          return false;

        val = found->second.value;
        return true;
      }

    private:

      struct TopicRecord
      {
        shared_ptr<IRtdPublisher> publisher;
        unordered_set<long> subscribers;
        shared_ptr<TValue> value;
      };

      unordered_map<wstring, TopicRecord> _records;

      list<pair<wstring, shared_ptr<TValue>>> _newValues;
      vector<pair<long, wstring>> _topicsToConnect;
      vector<long> _topicIdsToDisconnect;

      // Publishers which have been cancelled but haven't finished terminating
      list<shared_ptr<IRtdPublisher>> _cancelledPublishers;

      std::function<void()> _updateNotify;
      atomic<SAFEARRAY*> _readyUpdates;
      atomic<bool> _isRunning;

      // We use a separate lock for the newValues to avoid blocking too 
      // often: value updates are likely to come from other threads and 
      // simply need to write into newValues without accessing pub/sub info.
      // We use _lockRecords for all other synchronisation
      mutable mutex _mutexNewValues;
      mutable mutex _mutexNewSubscribers;
      mutable std::shared_mutex _lockRecords;

      std::thread _workerThread;
      std::condition_variable _workPendingNotifier;
      atomic<bool> _workPending = false;

      void notify() noexcept
      {
        _workPending = true;
        _workPendingNotifier.notify_one();
      }

      void setQuitFlag()
      {
        _isRunning = false;
      }

      bool isServerRunning() const
      {
        return _isRunning;
      }

      void workerThreadMain()
      {
        unordered_set<long> readyTopicIds;
        unordered_map<long, wstring> activeTopicIds;
        try
        {
          while (isServerRunning())
          {
            // The worker does all the work!  In this order
            //   1) Wait for wake notification
            //   2) Check if quit/stop has been sent
            //   3) Look for new values.
            //      a) If any, put the matching topicIds in readyTopicIds
            //      b) If Excel has picked up previous values, create an array of 
            //         updates and send an UpdateNotify.
            //   4) Run any topic connect requests
            //   5) Run any topic disconnect requests
            //   6) Repeat
            //
            decltype(_newValues) newValues;

            unique_lock lockValues(_mutexNewValues);
            // This slightly convoluted code protects against spurious wakes and 
            // 'lost' wakes, i.e. if the CV is signalled but the worker is not
            // in the waiting state.
            if (!_workPending)
              _workPendingNotifier.wait(lockValues, [&]() { return _workPending.load(); });
            _workPending = false;

            if (!isServerRunning())
              break;

            // Since _mutexNewValues is required to send updates, so we avoid holding it  
            // and quickly swap out the list of new values. 
            std::swap(newValues, _newValues);
            lockValues.unlock();

            if (!newValues.empty())
            {
              shared_lock lock(_lockRecords);
              auto iValue = newValues.begin();
              for (; iValue != newValues.end(); ++iValue)
              {
                auto record = _records.find(iValue->first);
                if (record == _records.end())
                  continue;
                record->second.value = iValue->second;
                readyTopicIds.insert(record->second.subscribers.begin(), record->second.subscribers.end());
              }
            }

            // When Excel calls RefreshData, it will take the SAFEARRAY in _readyUpdates and
            // atomically replace it with null. If this ptr is not null, we know Excel
            // has not yet picked up the new values, so we swap it out and resize the array
            // to include the latest ready topics. We issue another _updateNotify to Excel,
            // even if there are items in readyUpdates as sometimes things go out of sync and
            // Excel does not call RefreshData (exact reasons unknown).
            if (!readyTopicIds.empty())
            {
              const auto nReady = (ULONG)readyTopicIds.size();
              auto topicArray = _readyUpdates.exchange(nullptr);
              long nExisting = 0;

              if (topicArray)
              {
                SafeArrayGetUBound(topicArray, 2, &nExisting); 
                ++nExisting; // Bound is *inclusive*
                SAFEARRAYBOUND outer{ nExisting + nReady, 0 };
                SafeArrayRedim(topicArray, &outer);
              }
              else
              {
                SAFEARRAYBOUND bounds[] = { { 2u, 0 }, { nReady, 0 } };
                topicArray = SafeArrayCreate(VT_VARIANT, 2, bounds);
              }

              writeReadyTopicsArray(topicArray, readyTopicIds, nExisting);

              _readyUpdates.exchange(topicArray);

              _updateNotify();

              readyTopicIds.clear();
            }

            decltype(_topicsToConnect) topicsToConnect;
            decltype(_topicIdsToDisconnect) topicIdsToDisconnect;
            {
              unique_lock lock(_mutexNewSubscribers);

              std::swap(_topicIdsToDisconnect, topicIdsToDisconnect);
              std::swap(_topicsToConnect, topicsToConnect);
            }

            if (topicsToConnect.empty() && topicIdsToDisconnect.empty())
              continue;

            const bool doTrace = spdlog::default_logger_raw()->should_log(spdlog::level::trace);

            if (doTrace)
            {
              wstring msg = L"RTD connections: ";
              for (auto& x : topicsToConnect)
                msg.append(formatStr(L"(%s -> %d), ", x.second.c_str(), x.first));
              XLO_TRACE(msg);

              msg = L"RTD disconnections: ";
              for (auto& x : topicIdsToDisconnect)
                msg.append(formatStr(L"%d, ", x));
              XLO_TRACE(msg);
            }

            // The connect/disconnect logic seems unnecessarily tortuous. 
            // There are a few tricky bases to cover:
            //   1) Do not hold locks when calling functions on the publisher
            //      as they may re-enter the RTD framework.
            //   2) Catch any exceptions from publisher calls to avoid killing
            //      the RTD server
            //   3) We may have a disconnect and connect on the same publisher
            //      and want to avoid the publisher cancelling itself.
            //   3) We may have a disconnect and connect on the same topic id
            //      and need to ensure `activeTopicIds` ends in the correct
            //      state
            //   4) We may have a disconnect and connect on the same topic id
            //      *and* the same publisher. In this case we want to cancel
            //      out the requests. 

            // Sort pending connects / disconnects by topic ID, so we can more 
            // efficiently iterate one whilst looking up in the other. 
            std::sort(topicsToConnect.begin(), topicsToConnect.end(),
              [](auto& l, auto& r) { return l.first < r.first; });
            std::sort(topicIdsToDisconnect.begin(), topicIdsToDisconnect.end());

            // Lookup the topic names for the IDs to disconnect, then remove
            // them as they are no longer needed
            decltype(_topicsToConnect) topicsToDisconnect;
            if (!topicIdsToDisconnect.empty())
            {
              topicsToDisconnect.reserve(topicIdsToDisconnect.size());
              for (auto id : topicIdsToDisconnect)
              {
                auto p = activeTopicIds.find(id);
                if (p != activeTopicIds.end())
                {
                  topicsToDisconnect.emplace_back(id, p->second);
                  activeTopicIds.erase(p);
                }
                else
                  XLO_WARN("Could not find topic to disconnect for id {0}", id);
              }
            }

            // Step through the topics to connect, checking for a match
            // in topicsToDisconnect. If one is found, skip the connect
            // and zero the topic id which will skip the disconnect later.
            // Otherwise connect the topic and update the active IDs
            {
              auto iDisconnect = topicsToDisconnect.begin();
              const auto ttdEnd = topicsToDisconnect.end();
              for (auto& [topicId, topic] : topicsToConnect)
              {
                iDisconnect = std::lower_bound(iDisconnect, ttdEnd, topicId,
                  [](auto& l, auto& r) { return l.first < r; });
                if (iDisconnect != ttdEnd
                  && iDisconnect->first == topicId
                  && iDisconnect->second == topic)
                {
                  if (doTrace)
                    XLO_TRACE(L"RTD matched connect against disconnect for ({} -> {})", topicId, topic);
                  iDisconnect->first = 0;
                  ++iDisconnect;
                }
                else
                  connectTopic(topicId, topic);

                activeTopicIds.emplace(topicId, std::move(topic));
              }
            }

            // Now we can run the disconnects, skipping any zeroed which matched against a connection
            for (auto& [topicId, topic] : topicsToDisconnect)
              if (topicId != 0)
                disconnectTopic(topicId, topic);

            // Remove any cancelled publishers which have finalised.  Don't destroy 
            // them whilst holding the lock.  Note the i++ as the iterator passed to 
            // splice will be invalidated.
            decltype(_cancelledPublishers) finalisedPublishers;
            {
              unique_lock lock(_lockRecords);
              for (auto i = _cancelledPublishers.begin(); i != _cancelledPublishers.end(); ++i)
                if ((*i)->done())
                  finalisedPublishers.splice(finalisedPublishers.end(), _cancelledPublishers, i++);
            }

            // Now run the dtors, but catch to avoid killing the RTD server
            try 
            {
              finalisedPublishers.clear();
            }
            catch (const std::exception& e)
            {
              XLO_ERROR("Error during publisher desctructor: {}", e.what());
            }
          }
        }
        catch (const std::exception& e)
        {
          XLO_ERROR("RTD Worker thread exited with error: {}", e.what());
        }
        
        try
        {
          // Clear all records, destroy all publishers
          clear();
        }
        catch (const std::exception& e)
        {
          XLO_ERROR("Error during RTD worker cleanup: {}", e.what());
        }
      }

      // 
      // Creates a 2 x n safearray which has rows of:
      //     topicId | empty
      // With the topicId for each updated topic. The second column can be used
      // to pass an updated value to Excel, however, only string values are allowed
      // which is too restricive. Passing empty tells Excel to call the function
      // again to get the value
      //
      static void writeReadyTopicsArray(
        SAFEARRAY* data,
        const std::unordered_set<long>& topics,
        const long startRow = 0)
      {
        void* element = nullptr;
        auto iRow = startRow;
        for (auto topic : topics)
        {
          long index[] = { 0, iRow };
          auto ret = SafeArrayPtrOfIndex(data, index, &element);
          assert(S_OK == ret);
          *(VARIANT*)element = _variant_t(topic);

          index[0] = 1;
          ret = SafeArrayPtrOfIndex(data, index, &element);
          assert(S_OK == ret);
          *(VARIANT*)element = _variant_t();

          ++iRow;
        }
      }

      void connectTopic(long topicId, const wstring& topic)
      {
        // We need these values after we release the lock
        shared_ptr<IRtdPublisher> publisher;
        size_t numSubscribers;

        {
          XLO_TRACE(L"RTD: connecting '{}' to topicId '{}'", topic, topicId);
          unique_lock lock(_lockRecords);
          auto& record = _records[topic];
          publisher = record.publisher;
          record.subscribers.insert(topicId);
          numSubscribers = record.subscribers.size();
        }

        // Let the publisher know how many subscribers they now have.
        // We must not hold the lock when calling functions on the publisher
        // as they may try to call other functions on the RTD server. 
        if (publisher)
          publisher->connect(numSubscribers);
      }

      void disconnectTopic(long topicId, const wstring& topic)
      {
        shared_ptr<IRtdPublisher> publisher;
        size_t numSubscribers;
 
        try
        {
          // We must *not* hold the lock when calling methods of the publisher
          // as they may try to call other functions on the RTD server. So we
          // first handle the topic lookup and removing subscribers before
          // releasing the lock and notifying the publisher.
          {
            unique_lock lock(_lockRecords);
  
            auto& record = _records[topic];
            record.subscribers.erase(topicId);

            numSubscribers = record.subscribers.size();
            publisher = record.publisher;

            if (!publisher && numSubscribers == 0)
            {
              XLO_TRACE(L"Removing orphaned topic {}", topic);
              _records.erase(topic);
            }
          }

          if (!publisher)
            return;

          XLO_TRACE(L"RTD: disconnecting '{}' with topicId '{}'", topic, topicId);

          // If disconnect() causes the publisher to cancel its task,
          // it will return true here. We may not be able to just delete it: 
          // we have to wait until any threads it created have exited
          if (publisher->disconnect(numSubscribers))
          {
            const auto done = publisher->done();
            if (!done)
              publisher->stop();

            XLO_TRACE(L"Removing publisher and record for topic {}", topic);

            {
              unique_lock lock(_lockRecords);

              // Not done, so add to this list and check back later
              if (!done)
                _cancelledPublishers.emplace_back(publisher);

              // Disconnect should only return true when num_subscribers = 0, 
              // so it's safe to erase the entire record
              _records.erase(topic);
            }
          }
        }
        catch (const std::exception& e)
        {
          XLO_ERROR("RTD error whilst disconnecting '{}': {}", topicId, e.what());
        }
      }

      void clear()
      {
        // We must not hold any locks when calling functions on the publisher
        // as they may try to call other functions on the RTD server. 
        list<shared_ptr<IRtdPublisher>> publishers;
        {
          unique_lock lock(_lockRecords);

          for (auto& record : _records)
            if (record.second.publisher)
              publishers.emplace_back(std::move(record.second.publisher));

          _records.clear();
          _cancelledPublishers.clear();
        }

        for (auto& publisher : publishers)
        {
          try
          {
            publisher->disconnect(0);
            if (!publisher->done())
              publisher->stop();
          }
          catch (const std::exception& e)
          {
            XLO_INFO(L"Failed to stop producer: '{0}': {1}",
              publisher->topic(), utf8ToUtf16(e.what()));
          }
        }
      }
    };
  }
}
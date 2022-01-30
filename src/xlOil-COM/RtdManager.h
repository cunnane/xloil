#pragma once
#include <xloil/RtdServer.h>

struct tagSAFEARRAY;
using SAFEARRAY = tagSAFEARRAY;

namespace xloil
{
  namespace COM
  {
    /// <summary>
    /// Overridden Inteface for objects which handle calls from an
    /// RTD COM Server
    /// </summary>
    struct IRtdServerWorker
    {
      void start(std::function<void()>&& updateNotify);
      /// <summary>
      /// Ensure the given topic is being published and link it to the
      /// specified topicId
      /// </summary>
      void connect(long topicId, std::wstring&& topic);
      /// <summary>
      /// Delink the topicId (which may cause an associated topic publisher
      /// to be stopped)
      /// </summary>
      /// <param name="topicId"></param>
      void disconnect(long topicId);
      /// <summary>
      /// Returns 2 x n SafeArray which has rows of [topicId, empty] for
      /// each topicId with a new value
      /// </summary>
      SAFEARRAY* getUpdates();
      /// <summary>
      /// Stops any publishers. The server will be unavailable after
      /// this call.
      /// </summary>
      void quit();
    };

    template <class TValue>
    struct IRtdPublishManager
    {
      void update(std::wstring&& topic, const std::shared_ptr<TValue>& value);
      void addPublisher(const std::shared_ptr<IRtdPublisher>& job);
      bool dropPublisher(const wchar_t* topic);
      bool value(const wchar_t* topic, std::shared_ptr<const TValue>& val) const;
      void quit();
      /// <summary>
      /// Calls quit, then joins any publishers and worker threads. The object
      /// is ready for destruction after this call returns.
      /// </summary>
      void join();
    };

    std::shared_ptr<IRtdServer> newRtdServer(
      const wchar_t* progId, const wchar_t* clsid);
  }
}

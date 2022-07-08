#pragma once
#include <xloil/ExcelObj.h>
#include <xloil/PString.h>
#include <xloil/Events.h>
#include <xloil/Caller.h>
#include <xloil/Throw.h>
#include <unordered_map>
#include <unordered_set>
#include <string_view>
#include <mutex>

namespace xloil
{
  template<class T>
  struct CacheUniquifier
  {
    CacheUniquifier()
    {
      static wchar_t chr = L'\xC38';
      value = chr++;
    }
    wchar_t value;
  };

  template<wchar_t Value>
  struct CacheUniquifierIs
  {
    static constexpr wchar_t value = Value;
  };

  namespace detail
  {
    struct SharedPtrToPtr
    {
      template<class T>
      intptr_t operator()(const T& obj)
      {
        return (intptr_t)obj.get();
      }
    };
  }
  /// <summary>
  /// Creates a dictionary of TObj indexed by cell address.
  /// The cell address used is determined from the currently executing cell
  /// when the "add" method is called.
  /// 
  /// This class is used to allow data which we cannot or do not want to write
  /// to an Excel sheet to be passed between Excel user functions.
  /// 
  /// The cache add a single character uniquifier TUniquifier to returned 
  /// reference strings. This allows very fast rejection of invalid references
  /// during lookup. The uniquifier should be choosen to be unlikely to occur 
  /// at the start of a string before a '['.
  /// 
  /// Example
  /// -------
  /// <code>
  /// static ObjectCache<PyObject*>, L'&', false> theCache
  ///     = ObjectCache<PyObject*>, L'&', false > ();
  /// 
  /// ExcelObj refStr = theCache.add(new PyObject());
  /// PyObject* obj = theCache.fetch(refStr);
  /// </code>
  /// </summary>
  /// 
  template<class TObj, class TUniquifier, class ToPtr = detail::SharedPtrToPtr>
  class ObjectCache
  {
    template <class, class = void>
    struct DefaultTag
    {
      static constexpr std::wstring_view value = std::wstring_view();
    };

    template <class T>
    struct DefaultTag<T, std::void_t<decltype(T::tag)>>
    {
      static constexpr std::wstring_view value = T::tag;
    };

  public:
    using this_t = ObjectCache<TObj, TUniquifier, ToPtr>;
    static constexpr auto defaultTag = DefaultTag<TUniquifier>::value;

    struct Entry
    {
      TObj obj;
      CallerInfo caller;
    };

    struct Key
    {
      intptr_t ptr;
      uint32_t calcId;

      auto operator==(const Key& that) const
      {
        return calcId == that.calcId && ptr == that.ptr;
      }
    };

    struct KeyHash
    {
      size_t operator()(const Key& p) const noexcept
      {
        return boost_hash_combine(377, p.calcId, p.ptr);
      }
    };

  private:
    std::unordered_map<Key, Entry, KeyHash> _cache;

    mutable std::mutex _cacheLock;
    enum class Reaper { STOPPED = 0, PAUSED = 1, RUNNING = 2 };
    mutable std::atomic<Reaper> _reaperState;
    std::thread _reaper;
    std::condition_variable _reaperCycle;

    std::unordered_set<std::wstring> _closedWorkbooks;

    uint32_t _calcId;
    TUniquifier _uniquifier;

    std::shared_ptr<const void> _calcEndHandler;
    std::shared_ptr<const void> _workbookCloseHandler;
    std::shared_ptr<const void> _workbookOpenHandler;

    static constexpr uint8_t _PrefixLength = 2;
    static constexpr uint8_t _NumKeyBytes = sizeof(Key);
    static constexpr uint8_t _MinKeyLength = _NumKeyBytes + _PrefixLength;
    static constexpr auto _ReaperSleepTime = std::chrono::seconds(10);
    static constexpr wchar_t _CachePrefix = L'©';
    static constexpr uint16_t _CacheCharMask = 0x0100;
    
    // We save a comparison in valid() by comparing the first two wchars 
    // which must be "<prefix><uniquifier>" as a single uint32. Note the
    // endianness to get this the right way around.
    uint32_t _prefixInt = _CachePrefix + (uint32_t)(_uniquifier.value << 16);
    
    // Created via factory function
    ObjectCache()
      : _calcId(1)
    {
      _reaperState = Reaper::RUNNING;
      _reaper = std::thread(&this_t::reaperMain, this);
    }

  public:
    ~ObjectCache()
    {
      _reaperState = Reaper::STOPPED;
      _reaperCycle.notify_one();
      _reaper.join();
    }

    static auto create()
    {
      auto p = std::shared_ptr<this_t>(new this_t);

      p->_calcEndHandler =
        xloil::Event::AfterCalculate().weakBind(
          std::weak_ptr<this_t>(p), &this_t::onAfterCalculate);

      p->_workbookCloseHandler =
        xloil::Event::WorkbookAfterClose().weakBind(
          std::weak_ptr<this_t>(p), &this_t::onWorkbookClose);

      p->_workbookOpenHandler =
        xloil::Event::WorkbookOpen().weakBind(
          std::weak_ptr<this_t>(p), &this_t::onWorkbookOpen);

      return p;
    }

    bool valid(const std::wstring_view& cacheString) const
    {
      // We compare the first two wchars to a single int constructed aboves
      return cacheString.size() >= _MinKeyLength
        && *(decltype(_prefixInt)*)cacheString.data() == _prefixInt;
    }

    const TObj* fetch(const std::wstring_view& key) const
    {
      if (key.size() < _MinKeyLength)
        return nullptr;

      auto keyVal = strToKey(key);

      std::scoped_lock lock(std::adopt_lock, signalAndLock());

      const auto found = _cache.find(keyVal);
      return found == _cache.end()
        ? nullptr
        : &found->second.obj;
    }

    PString add(
      TObj&& obj,
      CallerInfo&& caller = CallerInfo(),
      const std::wstring_view& tag = defaultTag)
    {
      Entry entry{ std::move(obj), std::move(caller) };
      Key   key{ ToPtr()(entry.obj), _calcId };

      decltype(_cache)::iterator i;
      bool success;
      {
        std::scoped_lock lock(std::adopt_lock, signalAndLock());
        std::tie(i, success) = _cache.try_emplace(std::move(key), std::move(entry));

        assert(success);
        assert(i->first.ptr == ToPtr()(i->second.obj));
      }

      return keyToStr(i->first, tag);
    }

    auto reap()
    {
      std::unique_lock<std::mutex> lock(_cacheLock);
      lock.unlock();
      _reaperCycle.notify_one();
      Sleep(50);
      lock.lock();
      return _cache.size();
    }

    void onAfterCalculate()
    {
      // Called by Excel event so will always be synchonised
      ++_calcId; // Wraps at MAX_UINT - but this doesn't matter
    }

    void onWorkbookClose(const wchar_t* wbName)
    {
      // Called by Excel Event so will always be synchronised
      std::scoped_lock lock(std::adopt_lock, signalAndLock());
      _closedWorkbooks.insert(wbName);
    }

    void onWorkbookOpen(const wchar_t* /*wbPath*/, const wchar_t* wbName)
    {
      std::scoped_lock lock(std::adopt_lock, signalAndLock());
      _closedWorkbooks.erase(wbName);
    }

    auto begin() const { return _cache.cbegin(); }

    auto end() const { return _cache.cend(); }

    PString keyToStr(
      const Key& key,
      const std::wstring_view& tag = defaultTag) const
    {
      const auto tagLen = (uint8_t)tag.size();
      PString result(_MinKeyLength + tagLen);

      auto pStr = result.pstr();

      *pStr++ = _CachePrefix;
      *pStr++ = _uniquifier.value;

      if (tagLen > 0)
      {
        wmemcpy(pStr, tag.data(), tagLen);
        pStr += tagLen;
      }

      auto pKey = (const char*)&key;
      const auto pEnd = pKey + sizeof(Key);
      for (; pKey != pEnd; ++pStr, ++pKey)
        *pStr = (_CacheCharMask | *pKey);

      return std::move(result);
    }

  private:

    Key strToKey(const std::wstring_view& str) const
    {
      Key key;
      auto pStr = (char*)(str.data() + str.size() - _NumKeyBytes);

      auto pKey = (char*)&key;
      const auto pEnd = pKey + sizeof(Key);
      for (; pKey != pEnd; pStr += 2, ++pKey)
        *pKey = *pStr;

      return std::move(key);
    }

    auto& signalAndLock() const
    {
      _reaperState = Reaper::PAUSED;
      _cacheLock.lock();
      _reaperState = Reaper::RUNNING;
      return _cacheLock;
    }

    void reaperMain()
    {
      std::unordered_map<std::wstring, Key> lastCalcId;
      std::unordered_set<Key, KeyHash> keysToRemove;
      std::list<decltype(_cache)::node_type> nodesToDelete;

      while (_reaperState != Reaper::STOPPED)
      {
        nodesToDelete.clear();

        std::unique_lock<std::mutex> lock(_cacheLock);
        _reaperCycle.wait_for(lock, _ReaperSleepTime);

        if (_reaperState == Reaper::STOPPED)
          break;

        do
        {
          for (auto& k : keysToRemove)
            nodesToDelete.emplace_back(_cache.extract(k));

          keysToRemove.clear();

          for (auto& [key, val] : _cache)
          {
            if (_reaperState <= Reaper::PAUSED)
              break;

            // TODO: in C++20 use 'contains' and avoid temp string
            if (_closedWorkbooks.count(std::wstring(val.caller.workbook())))
            {
              keysToRemove.insert(key);
            }
            else
            {
              const auto address = val.caller.writeAddress();
              auto found = lastCalcId.find(address);
              if (found != lastCalcId.end())
              {
                auto& mostRecentKey = found->second;
                if (key.calcId < mostRecentKey.calcId)
                  keysToRemove.insert(key);
                else if (key.calcId > mostRecentKey.calcId)
                {
                  keysToRemove.insert(mostRecentKey);
                  lastCalcId.insert_or_assign(found, address, key);
                }
              }
              else
              {
                lastCalcId[address] = key;
                // TODO: lastCalcId.emplace(address, i->first.calcId);
              }
            }
          }

          // If we exited the above loop naturally, we've dealt with all the
          // closed workbooks
          if (_reaperState == Reaper::RUNNING)
            _closedWorkbooks.clear();

        } while (!keysToRemove.empty());
      }
    }
  };

  template<typename T>
  struct ObjectCacheFactory
  {
    static auto& cache() {
      static auto theInstance = ObjectCache<T, CacheUniquifier<T>>::create();
      return *theInstance;
    }
  };

  /// <summary>
  /// Constructs an object of type <typeparamref name="T"/> and adds
  /// it to a cache of identically typed objects. Returns the key
  /// string as an <see cref="ExcelObj"/>.  Essentially a wrapper
  /// around <code><![CDATA[cache.add(make_unique<T>(...))]]></code>.
  /// The cache is automatically constructed if it doesn't already
  /// exist.
  /// </summary>
  template<typename T, typename... Args>
  inline auto makeCached(Args&&... args)
  {
    return ObjectCacheFactory<std::unique_ptr<const T>>::cache().add(
      std::make_unique<T>(std::forward<Args>(args)...));
  }

  // TODO: consider abrogated caching where simple types are just returned un-cached
  template<typename T>
  inline auto addCached(
    const T* ptr,
    const std::wstring_view& name = std::wstring_view())
  {
    return ObjectCacheFactory<std::unique_ptr<const T>>::cache().add(
      std::unique_ptr<const T>(ptr), CallerInfo(), name);
  }

  /// <summary>
  /// Retrieves an object of type <typeparamref name="T"/> given its key.
  /// Returns nullptr if not found.
  /// </summary>
  template<typename T>
  inline const T* getCached(const std::wstring_view& key)
  {
    if (!ObjectCacheFactory<std::unique_ptr<const T>>::cache().valid(key))
      return nullptr;

    const auto* found = ObjectCacheFactory<std::unique_ptr<const T>>::cache().fetch(key);
    return found ? found->get() : nullptr;
  }
}
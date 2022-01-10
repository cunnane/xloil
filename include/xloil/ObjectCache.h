#pragma once
#include <xloil/ExcelObj.h>
#include <xloil/PString.h>
#include <xloil/Events.h>
#include <xloil/Caller.h>
#include <xloil/Throw.h>
#include <unordered_map>
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

  namespace detail
  {
    template<uint16_t NPadding>
    inline auto writeCacheId(
      const CallerLite& caller, const wchar_t* optionalName, size_t nameLength)
    {
      auto pstrLength = CallerLite::INTERNAL_REF_MAX_LEN + nameLength + NPadding + 1u;
      PString<> pascalStr((uint16_t)pstrLength);
      auto* buf = pascalStr.pstr() + 1;

      int nWritten = 1; // Leave space for uniquifier

      nWritten += caller.writeInternalAddress(buf, pstrLength - 1);
      // Check for a negative return condition from the above. This should not
      // be possible as we made the buffer larget than the max internal ref length
      assert(nWritten - 1 > 0 && nWritten <= CallerLite::INTERNAL_REF_MAX_LEN + 1);

      if (optionalName)
        wmemcpy_s(buf + nWritten - 1, pstrLength - nWritten, optionalName, nameLength);

      // Fix up length
      pascalStr.resize(uint16_t(nWritten + nameLength + NPadding));

      return pascalStr;
    }

    // We need to explicitly define our own lookup function for undordered_map
    // to use string_view objects without copying.  In std::map the find function 
    // is templated on the key type so we're just replicating that behaviour.
    template<class Val>
    struct Lookup : public std::unordered_map<std::wstring, Val>
    {
      using base = typename std::unordered_map<std::wstring, Val>;

      template <class T>
      _NODISCARD typename base::const_iterator search(const T& _Keyval) const
      {
        size_type _Bucket = std::hash<T>()(_Keyval) & _Mask;
        for (auto _Where = begin(_Bucket); _Where != end(_Bucket); ++_Where)
          if (_Where->first == _Keyval)
            return _Where;
        return (end());
      }
      template <class T>
      _NODISCARD typename base::iterator search(const T& _Keyval)
      {
        size_type _Bucket = std::hash<T>()(_Keyval) & _Mask;
        for (auto _Where = begin(_Bucket); _Where != end(_Bucket); ++_Where)
          if (_Where->first == _Keyval)
            return _Where;
        return (end());
      }
    };

    template<typename TObj>
    class CellCache
    {
    private:
      size_t _calcId;
      std::vector<TObj> _objects;
      TObj _obj;

    public:
      CellCache(TObj&& obj, size_t calcId)
        : _calcId(calcId)
        , _obj(std::move(obj))
      {}

      void getStaleObjects(size_t calcId, std::vector<TObj>& stale)
      {
        if (_calcId != calcId)
        {
          _objects.swap(stale);
          stale.emplace_back(std::move(_obj));
        }
      }

      size_t count() const { return _objects.size() + 1; }

      size_t add(TObj&& obj, size_t calcId)
      {
        if (_calcId != calcId)
        {
          std::swap(_obj, obj);
          _calcId = calcId;
          _objects.clear();
        }
        else
          _objects.emplace_back(std::forward<TObj>(obj));
        return _objects.size();
      }

      const TObj* fetch(size_t i) const
      {
        if (i == 0)
          return &_obj;
        else if (i <= _objects.size())
          return &_objects[i - 1];
        else
          return nullptr;
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
  template<class TObj, class TUniquifier, bool TReverseLookup = false>
  class ObjectCache
  {
  private:
    typedef ObjectCache<TObj, TUniquifier, TReverseLookup> self;
    typedef detail::CellCache<TObj> CellCache;

    detail::Lookup<CellCache> _cache;
    mutable std::mutex _cacheLock;

    size_t _calcId;

    struct Reverse
    {
      std::unordered_map<const TObj*, std::wstring> map;
      mutable std::mutex lock;
    };
    typename std::conditional<TReverseLookup, Reverse, char>::type _reverseLookup;

    std::shared_ptr<const void> _calcEndHandler;
    std::shared_ptr<const void> _workbookCloseHandler;

    void onAfterCalculate()
    {
      // Called by Excel event so will always be synchonised
      ++_calcId; // Wraps at MAX_UINT - but this doesn't matter
    }

    /// <summary>
    /// Used to append cell count to end of reference
    /// </summary>
    static constexpr uint8_t PADDING = 2;

    TUniquifier _uniquifier;

    ObjectCache()
      : _calcId(1)
    {}

  public:

    /// <summary>
    /// Max length of a cache key. Because keys are derived from CallerLite using internal 
    /// references or otherwise, this length can be guaranteed
    /// </summary>
    static constexpr uint16_t KEY_MAX_LEN = 1 + CallerLite::INTERNAL_REF_MAX_LEN + PADDING;

    static auto create(bool reapOnWorkbookClose = true)
    {
      auto p = std::shared_ptr<ObjectCache>(new ObjectCache);
      p->_calcEndHandler =
        xloil::Event::AfterCalculate().weakBind(std::weak_ptr<ObjectCache>(p), &self::onAfterCalculate);

      if (reapOnWorkbookClose)
        p->_workbookCloseHandler =
        xloil::Event::WorkbookAfterClose().weakBind(std::weak_ptr<ObjectCache>(p), &self::onWorkbookClose);

      return p;
    }

    const TObj* fetch(const std::wstring_view& key) const
    {
      if (key.size() < PADDING) return nullptr;

      const auto iResult = readCount(key[key.size() - 1]);
      const auto cacheKey = key.substr(0, key.size() - PADDING);

      std::scoped_lock lock(_cacheLock);
      const auto found = _cache.search(cacheKey);

      return found == _cache.end()
        ? nullptr
        : found->second.fetch(iResult);
    }

    ExcelObj add(
      TObj&& obj,
      const CallerLite& caller = CallerLite(),
      const wchar_t* name = nullptr,
      const size_t nameLength = 0
    )
    {
      auto fullKey = detail::writeCacheId<PADDING>(caller, name, nameLength);
      fullKey[0] = _uniquifier.value;

      auto cacheKey = fullKey.view(0, fullKey.length() - PADDING);

      typename std::conditional<TReverseLookup, std::vector<TObj>, char>::type staleObjects;
      decltype(_cache)::iterator found;

      uint8_t iPos = 0;
      {
        std::scoped_lock lock(_cacheLock);

        found = _cache.search(cacheKey);
        if (found == _cache.end())
        {
          found = _cache.emplace(
            std::make_pair(
              std::wstring(cacheKey),
              CellCache(std::forward<TObj>(obj), _calcId))).first;
        }
        else
        {
          if constexpr (TReverseLookup)
            found->second.getStaleObjects(_calcId, staleObjects);

          iPos = (uint8_t)found->second.add(std::forward<TObj>(obj), _calcId);
        }
      }

      writeCount(fullKey.end() - PADDING, iPos);

      if constexpr (TReverseLookup)
      {
        std::scoped_lock lock(_reverseLookup.lock);
        for (auto& x : staleObjects)
          _reverseLookup.map.erase(&x);
        _reverseLookup.map.insert(std::make_pair(
          found->second.fetch(iPos),
          fullKey.string()));
      }

      return ExcelObj(std::move(fullKey));
    }

    /// <summary>
    /// Remove the given cache reference and any associated objects
    /// This should only be called with manually specifed cache reference
    /// strings. Note the counter (,NNN) after the cache reference is ignored
    /// if specifed and all matching objects are removed.
    /// </summary>
    /// <param name="key">cache reference to remove</param>
    /// <returns>true if removal succeeded, otherwise false</returns>
    bool erase(const std::wstring_view& key)
    {
      auto cacheKey = key.substr(0, key.length() - PADDING);

      std::scoped_lock lock(_cacheLock);
      auto found = _cache.search(cacheKey);
      if (found == _cache.end())
        return false;
      _cache.erase(found);
      return true;
    }

    void onWorkbookClose(const wchar_t* wbName)
    {
      // Called by Excel Event so will always be synchonised
      const auto len = wcslen(wbName);
      auto i = _cache.begin();
      while (i != _cache.end())
      {
        // Key looks like [WbName]BlahBlah - check for match
        if (wcsncmp(wbName, i->first.c_str() + 1, len) == 0)
        {
          if constexpr (TReverseLookup)
          {
            auto& cellCache = i->second;
            for (size_t k = 0; k < cellCache.count(); ++k)
              _reverseLookup.map.erase(cellCache.fetch(k));
          }
          i = _cache.erase(i);
        }
        else
          ++i;
      }
    }

    auto begin() const
    {
      return _cache.cbegin();
    }

    auto end() const
    {
      return _cache.cend();
    }

    std::wstring writeKey(
      const std::wstring_view& cacheKey,
      size_t count) const
    {
      std::wstring key;
      key.resize(cacheKey.length() + PADDING);
      key = cacheKey;
      writeCount(key.data() + cacheKey.length(), count);
      return key;
    }

    bool valid(const std::wstring_view& cacheString)
    {
      return cacheString.size() > 4
        && cacheString[0] == _uniquifier.value
        && cacheString[1] == L'['
        && cacheString[cacheString.length() - PADDING] == L',';
    }

    template<bool B = TReverseLookup>
    std::enable_if_t<B, const std::wstring*>
      findKey(const TObj* obj) const
    {
      if constexpr (TReverseLookup)
      {
        auto found = _reverseLookup.map.find(obj);
        return found == _reverseLookup.map.end() ? nullptr : &found->second;
      }
    }

  private:

    size_t readCount(wchar_t count) const
    {
      return (size_t)(count - 65);
    }

    /// Add cell object count in form ",X"
    void writeCount(wchar_t* key, size_t iPos) const
    {
      key[0] = L',';
      // An offset of 65 means we start with 'A'
      key[1] = (wchar_t)(iPos + 65);
    }
  };

  template<typename T>
  struct ObjectCacheFactory
  {
    static auto& cache() {
      static auto theInstance = ObjectCache<T, CacheUniquifier<T>, false>::create();
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
    const wchar_t* name = nullptr,
    const size_t nameLength = 0)
  {
    return ObjectCacheFactory<std::unique_ptr<const T>>::cache().add(
      std::unique_ptr<const T>(ptr), CallerLite(), name, nameLength);
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
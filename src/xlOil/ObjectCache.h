#pragma once
#include "ExcelCall.h"
#include "ExcelObj.h"
#include "Events.h"
#include "ExcelState.h"
#include "EntryPoint.h"
#include "Interface.h"
#include <unordered_map>
#include <string_view>
#include <mutex>

namespace xloil
{
  namespace impl
  {
    inline PString<> writeCacheId(
      wchar_t uniquifier, const CallerInfo& caller, uint16_t padding)
    {
      PString<> pascalStr(caller.fullAddressLength() + padding + 1);
      auto* buf = pascalStr.pstr();

      wchar_t nWritten = 0;
      // Cache uniquifier character
      *(buf++) = uniquifier;
      ++nWritten;

      // Full cell address: eg. [wbName]wsName!RxCy
      nWritten += (wchar_t)caller.writeFullAddress(buf, pascalStr.length() - 1);

      // Fix up length
      pascalStr.resize(nWritten + padding);

      return pascalStr;
    }
  }

  template<class TObj, wchar_t TUniquifier, bool TReverseLookup = false>
  class ObjectCache
  {
  private:
    typedef ObjectCache<TObj, TUniquifier, TReverseLookup> self;
    class CellCache
    {
    private:
      size_t _calcId;
      std::vector<TObj> objects;

    public:
      CellCache() : _calcId(0) {}

      bool getStaleObjects(size_t calcId, std::vector<TObj>& stale)
      {
        if (_calcId != calcId)
        {
          _calcId = calcId;
          objects.swap(stale);
          return true;
        }
        return false;
      }

      size_t add(TObj&& obj)
      {
        objects.emplace_back(std::forward<TObj>(obj));
        return objects.size() - 1;
      }

      bool tryFetch(size_t i, TObj& obj)
      {
        if (i >= objects.size())
          return false;
        obj = objects[i];
        return true;
      }

    };

  private:
    template <class Val> 
    using Lookup = std::unordered_map<std::wstring, std::shared_ptr<Val>>;
    using WorkbookCache = Lookup<CellCache>;

    Lookup<WorkbookCache> _cache;
    mutable std::mutex _cacheLock;

    size_t _calcId;

    typename std::conditional<TReverseLookup, 
      std::unordered_map<TObj, std::wstring>, 
      int>::type  _reverseLookup;
    typename std::conditional<TReverseLookup,
      std::mutex, 
      int>::type _reverseLookupLock;

    std::shared_ptr<const void> _calcEndHandler;
    std::shared_ptr<const void> _workbookCloseHandler;

    void expireObjects()
    {
      // Called by Excel event so will always be synchonised
      ++_calcId; // Wraps at MAX_UINT- doesn't matter
    }

    size_t addToCell(TObj&& obj, CellCache& cacheVal, std::vector<TObj>& staleObjects)
    {
      cacheVal.getStaleObjects(_calcId, staleObjects);
      return cacheVal.add(std::forward<TObj>(obj));
    }

    template<class V> V& findOrAdd(Lookup<V>& m, std::wstring_view keyView)
    {
      // TODO: YUK! Fix with boost?
      std::wstring key(keyView);
      auto found = m.find(key);
      if (found == m.end())
      {
        auto p = std::make_shared<V>();
        m.insert(std::make_pair(std::wstring(key), p));
        return *p;
      }
      return *found->second;
    }

    template<class V> V* find(Lookup<V>& m, std::wstring_view keyView)
    {
      std::wstring key(keyView);
      auto found = m.find(key);
      if (found == m.end())
        return nullptr;
      return found->second.get();
    }

  public:
    ObjectCache()
      : _calcId(1)
    {
      using namespace std::placeholders;

      _calcEndHandler = std::static_pointer_cast<const void>(
        xloil::Event::AfterCalculate().bind(std::bind(std::mem_fn(&self::expireObjects), this)));
      
      _workbookCloseHandler = std::static_pointer_cast<const void>(
        xloil::Event::WorkbookAfterClose().bind(std::bind(std::mem_fn(&self::onWorkbookClose), this, _1)));
    }

    bool fetch(const std::wstring_view& cacheString, TObj& obj)
    {
      if (cacheString[0] != TUniquifier || cacheString[1] != L'[')
        return false;

      constexpr auto npos = std::wstring_view::npos;

      const auto firstBracket = 1;
      const auto lastBracket = cacheString.find_last_of(']');
      if (lastBracket == npos)
        return false;
      const auto comma = cacheString.find_first_of(',', lastBracket);

      auto workbook = cacheString.substr(firstBracket + 1, lastBracket - firstBracket - 1);
      auto sheetRef = cacheString.substr(lastBracket + 1,
        comma == npos ? npos : comma - lastBracket - 1);

      int iResult = 0;
      if (comma != npos)
      {
        // Oh dear, there's no std::from_chars for wchar_t
        wchar_t tmp[4];
        wcsncpy_s(tmp, 4, cacheString.data() + comma + 1, cacheString.length() - comma - 1);
        iResult = _wtoi(tmp);
      }

      {
        std::scoped_lock lock(_cacheLock);

        auto* wbCache = find(_cache, workbook);
        if (!wbCache)
          return false;

        auto* cellCache = find(*wbCache, sheetRef);
        if (!cellCache)
          return false;

        return cellCache->tryFetch(iResult, obj);
      }
    }

    ExcelObj add(TObj&& obj)
    {
      CallerInfo caller;
      constexpr unsigned char padding = 5;

      auto key = impl::writeCacheId(TUniquifier, caller, padding);

      // Capture workbook name. pascalStr should have X[wbName]wsName!cellRef.
      // Search backwards because wbName may contain ']'
      auto lastBracket = key.rchr(L']');
      auto wbName = std::wstring_view(key.pstr() + 2, lastBracket - 2);

      // Capture sheet ref, i.e. wsName!cellRef
      // Can use wcslen here because of the null padding
      auto wsRef = std::wstring_view(key.pstr() + lastBracket + 1,
        key.length() - padding - lastBracket - 1);

      std::vector<TObj> staleObjects;
      size_t iPos = 0;
      {
        std::scoped_lock lock(_cacheLock);

        auto& cellCache = fetchOrAddCell(wbName, wsRef);
        iPos = addToCell(std::forward<TObj>(obj), cellCache, staleObjects);
      }

      unsigned char nPaddingUsed = 0;
      // Add cell object count in form ",XXX"
      if (iPos > 0)
      {
        auto buf = const_cast<wchar_t*>(key.end()) - padding;
        *(buf++) = L',';
        // TODO: this will fail if iPos > 999
        _itow_s(int(iPos), buf, padding - 1, 10);
        nPaddingUsed = 1 + (decltype(nPaddingUsed))wcsnlen(buf, padding - 1);
      }
        
      key.resize(key.length() - padding + nPaddingUsed);

      if constexpr (TReverseLookup)
      {
        std::scoped_lock lock(_reverseLookupLock);
        for (auto o : staleObjects)
          _reverseLookup.erase(o);
        _reverseLookup.insert(std::make_pair(obj, key.string()));
      }

      return ExcelObj(std::move(key));
    }

  private:
    CellCache& fetchOrAddCell(const std::wstring_view& wbName, const std::wstring_view& wsRef)
    {
      auto& wbCache = findOrAdd(_cache, wbName);
      return findOrAdd(wbCache, wsRef);
    }

    void onWorkbookClose(const wchar_t* wbName)
    {
      // Called by Excel Event so will always be synchonised
      if constexpr (TReverseLookup)
      {
        auto found = _cache.find(wbName);
        if (found != _cache.end())
        {
          for (auto& cell : *found->second)
            for (auto& obj : cell.second->objects)
              _reverseLookup.erase(obj);
        }
      }
      _cache.erase(wbName);
    }
  };
}
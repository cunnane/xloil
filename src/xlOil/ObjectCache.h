#pragma once
#include "ExcelCall.h"
#include "ExcelObj.h"
#include "Events.h"
#include "ExcelState.h"
#include "EntryPoint.h"
#include "Interface.h"
#include <regex>
#include <unordered_map>
#include <string_view>
#include <mutex>

namespace xloil
{
  namespace impl
  {
    inline wchar_t* writeCacheId(wchar_t uniquifier, const CallerInfo& caller, int padding)
    {
      size_t totalLen = caller.fullAddressLength() + 1 + padding;
      wchar_t* pascalStr = new wchar_t[totalLen + 1];
      auto* buf = pascalStr + 1;

      // Cache uniquifier character
      *(buf++) = uniquifier;
      --totalLen;

      // Full cell address: eg. [wbName]wsName!RxCy
      auto addressLen = caller.writeFullAddress(buf, totalLen);
      buf += addressLen;

      // Pad with nulls
      wmemset(buf, L'\0', padding);

      // Fix up length
      pascalStr[0] = wchar_t(addressLen + 1 + padding);

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

    const std::wregex _cacheRefMatcher;
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
      : _cacheRefMatcher(LR"(\[([^\]]+)\]([^,]+),?(\d*).*)", std::regex_constants::optimize)
      , _calcId(1)
    {
      using namespace std::placeholders;

      _calcEndHandler = std::static_pointer_cast<const void>(
        xloil::Event_CalcEnded().bind(std::bind(std::mem_fn(&self::expireObjects), this)));
      
      _workbookCloseHandler = std::static_pointer_cast<const void>(
        xloil::Event_WorkbookClose().bind(std::bind(std::mem_fn(&self::onWorkbookClose), this, _1)));
    }

    bool fetch(const wchar_t* str, size_t length, TObj& obj)
    {
      std::wstring nastyTemp(str, str + length);
      return fetch(nastyTemp.c_str(), obj);
    }

    bool fetch(const wchar_t* cacheString, TObj& obj)
    {
      // TODO: write function without using regex, surely it would be quicker?
      // And we can remove nastytemp!

      if (cacheString[0] != TUniquifier || cacheString[1] != L'[')
        return false;

      std::wcmatch match;
      std::regex_match(cacheString + 1, match, _cacheRefMatcher);
      if (match.size() != 4)
        return false;

      auto workbook = match[1].str();
      auto sheetRef = match[2].str();
      auto position = match[3].str();
      size_t iResult = position.empty() ? 0 : _wtoi64(position.c_str());

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
      constexpr int padding = 4;

      auto* pascalStr = impl::writeCacheId(TUniquifier, caller, padding);

      // Oh the things we do to avoid a string copy
      auto len = pascalStr[0];
      auto* str = pascalStr + 1;

      // Capture workbook name.  We should have X[wbName]wsName!cellRef
      auto lastBracket = PString<>::wmemrchr(str + len, L']', len);
      auto wbName = std::wstring_view(str + 2, lastBracket - str - 2);

      // Capture sheet ref.  
      // Can use wcslen here because of the null padding
      auto wsRef = std::wstring_view(lastBracket + 1, len - (lastBracket - str) - 1 - padding);

      std::vector<TObj> staleObjects;
      size_t iPos = 0;
      {
        std::scoped_lock lock(_cacheLock);

        auto& cellCache = fetchOrAddCell(wbName, wsRef);
        iPos = addToCell(std::forward<TObj>(obj), cellCache, staleObjects);
      }

      // Add cell object count in form ",XXX"
      if (iPos > 0)
      {
        auto pStr = pascalStr + 1 + len - padding;
        *(pStr++) = L',';
        // TODO: this will fail if iPos > 999
        _itow_s(int(iPos), pStr, padding - 1, 10);
      }

      auto key = PString<>::own(pascalStr);

      if constexpr (TReverseLookup)
      {
        std::scoped_lock lock(_reverseLookupLock);
        for (auto o : staleObjects)
          _reverseLookup.erase(o);
        _reverseLookup.insert(std::make_pair(obj, key.string()));
      }

      return ExcelObj(key);
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
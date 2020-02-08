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

namespace xloil
{
  template<class TObj, bool TReverseLookup = false>
  class ObjectCache
  {
  private:
    typedef ObjectCache<TObj, TReverseLookup> self;
    class CellCache
    {
    private:
      size_t _calcId;

    public:
      std::vector<TObj> objects;

      CellCache() : _calcId(0) {}

      void expireAll() { _calcId = 0; }

      bool removeExpired(size_t calcId)
      {
        if (_calcId != calcId)
        {
          _calcId = calcId;
          return true;
        }
        return false;
      }

      size_t add(const TObj& obj)
      {
        objects.emplace_back(obj);
        return objects.size() - 1;
      }
    };

  private:
    template <class Val> using Lookup = std::unordered_map<std::wstring, std::shared_ptr<Val>>;
    typedef Lookup<CellCache> WorkbookCache;

    std::wregex _cacheRefMatcher;
    wchar_t _uniquifier;
    Lookup<WorkbookCache> _cache;
    typename std::conditional<TReverseLookup, std::unordered_map<TObj, std::wstring>, int>::type  _reverseLookup;
    size_t _calcId;
    std::shared_ptr<const void> _calcEndHandler;
    std::shared_ptr<const void> _workbookCloseHandler;

    void expireObjects()
    {
      ++_calcId; // Wraps at MAX_UINT- doesn't matter
    }

    size_t addToCell(const TObj& obj, CellCache& cacheVal)
    {
      if (cacheVal.removeExpired(_calcId))
      {
        if constexpr (TReverseLookup)
        {
          for (auto o : cacheVal.objects)
            _reverseLookup.erase(o);
        }
        cacheVal.objects.clear();
      }

      return cacheVal.add(obj);
    }

    template<class V> V& findOrAdd(Lookup<V>& m, std::wstring_view keyView)
    {
      // TODO: YUK! Fix with boost
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
    ObjectCache(wchar_t cacheUniquifier)
      : _cacheRefMatcher(LR"(\[([^\]]+)\]([^,]+),?(\d*).*)", std::regex_constants::optimize)
      , _uniquifier(cacheUniquifier)
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

      if (cacheString[0] != _uniquifier || cacheString[1] != L'[')
        return false;

      std::wcmatch match;
      std::regex_match(cacheString + 1, match, _cacheRefMatcher);
      if (match.size() != 4)
        return false;

      auto workbook = match[1].str();
      auto sheetRef = match[2].str();
      auto position = match[3].str();
      size_t iResult = position.empty() ? 0 : _wtoi64(position.c_str());

      auto* wbCache = find(_cache, workbook);
      if (!wbCache)
        return false;

      auto* cellCache = find(*wbCache, sheetRef);
      if (!cellCache)
        return false;

      if (cellCache->objects.size() <= iResult)
        return false;

      obj = cellCache->objects[iResult];
      return true;
    }

    // Like wmemchr but backwards!
    const wchar_t* wmemrchr(const wchar_t* ptr, wchar_t wc, size_t num)
    {
      for (; num; --ptr, --num)
        if (*ptr == wc)
          return ptr;
      return nullptr;
    }

    ExcelObj add(const TObj& obj)
    {
      CallerInfo caller;
      const int padding = 4;

      auto* pascalStr = writeCacheId(_uniquifier, caller, padding);

      // Oh the things we do to avoid a string copy
      size_t len = pascalStr[0];
      wchar_t* str = pascalStr + 1;

      // Capture workbook name.  We should have X[wbName]wsName!cellRef
      auto lastBracket = wmemrchr(str + len, L']', len);
      auto wbName = std::wstring_view(str + 2, lastBracket - str - 2);

      // Capture sheet ref.  
      // Can use wcslen here because of the null padding
      //assert(wcslen(lastBracket) == len - (lastBracket - str) - padding);
      auto wsRef = std::wstring_view(lastBracket + 1, len - (lastBracket - str) - 1 - padding);

      auto& cellCache = fetchOrAddCell(wbName, wsRef);
      auto iPos = addToCell(obj, cellCache);

      // Add cell object count in form ",XXX"
      if (iPos > 0)
      {
        auto pStr = pascalStr + 1 + len - padding;
        *(pStr++) = L',';
        // TODO: this will fail if iPos > 999
        _itow_s(int(iPos), pStr, padding - 1, 10);
      }

      auto key = PString(pascalStr);
      if constexpr (TReverseLookup)
        _reverseLookup.insert(std::make_pair(obj, key.string()));

      return ExcelObj(key);
    }

    CellCache& fetchOrAddCell(const std::wstring_view& wbName, const std::wstring_view& wsRef)
    {
      auto& wbCache = findOrAdd(_cache, wbName);
      return findOrAdd(wbCache, wsRef);
    }

    static wchar_t* writeCacheId(wchar_t uniquifier, const CallerInfo& caller, int padding)
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
      // return ret.release();
    }

    void onWorkbookClose(const wchar_t* wbName)
    {
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
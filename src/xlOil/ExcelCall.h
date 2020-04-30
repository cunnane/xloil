#pragma once
#include "ExcelObj.h"
#include "Throw.h"
#include <vector>
#include <memory>
#include <list>

namespace xloil
{
  template <class TIterable> struct Splatter
  {
    Splatter(const TIterable& iter) : _obj(iter) {}
    const TIterable& operator()() const { return _obj; }
    const TIterable& _obj;
  };

  // Another option to this would be to detect iterable with sfinae but
  // then is splatting the iterable more natural than converting it to 
  // an array?
  template <class TIterable>
  auto unpack(const TIterable& iterable)
  {
    return Splatter<TIterable>(iterable);
  }

  XLOIL_EXPORT int callExcelRaw(int func, ExcelObj* result, int nArgs = 0, const ExcelObj** args = nullptr);
  inline int callExcelRaw(int func, ExcelObj* result, const ExcelObj* arg)
  {
    auto p = arg;
    return callExcelRaw(func, result, 1, &p);
  }

  XLOIL_EXPORT const wchar_t* xlRetCodeToString(int xlret);

  namespace detail
  {
    class CallArgHolder
    {
    private:
      std::list<ExcelObj> _temporary;
      std::vector<const ExcelObj*> _argVec;

    public:
      template<class... Args> CallArgHolder(Args&&... args)
      {
        _argVec.reserve(sizeof...(args));
        add(std::forward<Args>(args)...);
      }

      const ExcelObj** ptrToArgs()
      {
        return (&_argVec[0]);
      }

      size_t nArgs() const { return _argVec.size(); }
 
      template<class T> void add(const T& arg)
      {
        _temporary.emplace_back(arg);
        _argVec.push_back(&_temporary.back());
      }
      void add(const ExcelObj& arg)
      {
        _argVec.push_back(&arg);
      }
      void add(const XLOIL_XLOPER* arg)
      {
        if (arg)
          _argVec.push_back((const ExcelObj*)arg);
        else
          add<nullptr_t>(nullptr);
      }

      template <class TIter>
      void add(Splatter<TIter>&& splatter)
      {
        for (const auto& x : splatter())
          add(x);
      }
      template<class T> void add(T&& arg)
      {
        _temporary.emplace_back(arg);
        _argVec.push_back(&_temporary.back());
      }

      template<class T, class...Args>
      void add(const T& first, Args&&... theRest)
      {
        add(first);
        add(std::forward<Args>(theRest)...);
      }
    };
  }

  template<typename... Args>
  inline ExcelObj callExcel(int func, Args&&... args)
  {
    auto[result, ret] = tryCallExcel(func, std::forward<Args>(args)...);
    switch (ret)
    {
    case msxll::xlretSuccess:
      break;
    case msxll::xlretAbort:
      throw new ExcelAbort();
    default:
      XLO_THROW(L"Call to Excel failed: {0}", xlRetCodeToString(ret));
    }
    return std::forward<ExcelObj>(result);
  }

  template<typename... Args>
  inline std::pair<ExcelObj, int> 
    tryCallExcel(int func, Args&&... args) noexcept
  {
    auto result = std::make_pair(ExcelObj(), 0);
    detail::CallArgHolder holder(std::forward<Args>(args)...);
    result.second = callExcelRaw(func, &result.first, holder.nArgs(), holder.ptrToArgs());
    result.first.fromExcel();
    return std::forward<std::pair<ExcelObj, int>>(result);
  }

  inline std::pair<ExcelObj, int> 
    tryCallExcel(int func)
  {
    auto result = std::make_pair(ExcelObj(), 0);
    result.second = callExcelRaw(func, &result.first, 0, 0);
    result.first.fromExcel();
    return result;
  }

  inline std::pair<ExcelObj, int> 
    tryCallExcel(int func, const ExcelObj& arg)
  {
    auto result = std::make_pair(ExcelObj(), 0);
    auto p = &arg;
    result.second = callExcelRaw(func, &result.first, 1, &p);
    result.first.fromExcel();
    return result;
  }

  class ExcelAbort : public std::runtime_error
  {
  public:
    ExcelAbort() : std::runtime_error("Excel abort called") {}
  };
}


#pragma once
#include "ExcelObj.h"
#include "Log.h"
#include <vector>
#include <memory>


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

  template<class TTarget, class TTemp, class TFirst>
  void appendVector(std::vector<TTarget>& v, std::list<TTemp>& tmp, const TFirst& first)
  {
    tmp.emplace_back(first);
    appendVector(v, tmp, tmp.back().cptr());
  }
  template<class TTarget, class TTemp>
  void appendVector(std::vector<TTarget>& v, std::list<TTemp>& tmp, const nullptr_t& first)
  {
    // TODO: Missing type singleton?
    tmp.emplace_back(nullptr);
    appendVector(v, tmp, tmp.back().cptr());
  }
  template<class TTarget, class TTemp>
  void appendVector(std::vector<TTarget>& v, std::list<TTemp>& tmp, const ExcelObj& first)
  {
    v.push_back(&first);
  }
  template<class TTarget, class TTemp>
  void appendVector(std::vector<TTarget>& v, std::list<TTemp>& tmp, const ExcelObj* first)
  {
    v.push_back(first);
  }

  template<class TTarget, class TTemp>
  void appendVector(std::vector<TTarget>& v, std::list<TTemp>& tmp, const XLOIL_XLOPER* first)
  {
    v.push_back((const ExcelObj*) first);
  }

  template<class TTarget, class TTemp, class TIter>
  void appendVector(std::vector<TTarget>& v, std::list<TTemp>& tmp, const Splatter<TIter>& first)
  {
    for (const auto& x : first())
      appendVector(v, tmp, x);
  }

  template<class TTarget, class TTemp, class T, class...Args>
  void appendVector(std::vector<TTarget>& v, std::list<TTemp>& tmp, const T& first, Args&&... theRest)
  {
    appendVector(v, tmp, first);
    appendVector(v, tmp, theRest...);
  }


  // TODO: use this rather
  class ArgHolder
  {
  private:
    std::vector<std::shared_ptr<ExcelObj>> _temporary;
    std::vector<const ExcelObj*> _argVec;

  public:

    template<class... Args> ArgHolder(const Args&&... args)
    {
      _temporary.reserve(sizeof...(args));
      _argVec.reserve(sizeof...(args));
      add(args...);
    }

    const ExcelObj** ptrToArgs()
    {
      return (&_argVec[0]);
    }

    template<class T> void add(const T& first)
    {
      auto p = std::make_shared<ExcelObj>(first);
      _temporary.push_back(p);
      add(p.cptr());
    }

    // TODO: do we need this separatly?
    template<> void add(const nullptr_t& first)
    {
      auto p = std::make_shared<ExcelObj>(nullptr);
      _temporary.push_back(p);
      add(p->cptr());
    }
    template<> void add(const ExcelObj& first)
    {
      _argVec.push_back(&first);
    }
    void add(const XLOIL_XLOPER* first)
    {
      _argVec.push_back((const ExcelObj*)first);
    }

    template<class T, class...Args>
    void add(const T& first, const Args&... theRest)
    {
      add(first);
      add(theRest...);
    }
  };

  inline const ExcelObj** toArgPtr(std::vector<const ExcelObj*>& args)
  {
    return (&args[0]);
  }

  template<typename... Args>
  inline ExcelObj callExcel(int func, Args&&... args)
  {
    auto[result, ret] = tryCallExcel(func, std::forward<Args>(args)...);
    if (ret != msxll::xlretSuccess)
      XLO_THROW(L"Call to Excel failed: {0}", xlRetCodeToString(ret));
    return result;
  }

  template<typename... Args>
  inline std::pair<ExcelObj, int> tryCallExcel(int func, Args&&... args) noexcept
  {
    auto result = std::make_pair(ExcelObj(), 0);
    std::list<ExcelObj> tmpVec; // TODO: memory pool like in frmwrk is better?
    std::vector<const ExcelObj*> argVec;
    appendVector(argVec, tmpVec, std::forward<Args>(args)...);
    result.second = callExcelRaw(func, &result.first, int(argVec.size()), toArgPtr(argVec));
    result.first.fromExcel();
    return result;
  }

  inline std::pair<ExcelObj, int> tryCallExcel(int func)
  {
    auto result = std::make_pair(ExcelObj(), 0);
    result.second = callExcelRaw(func, &result.first, 0, 0);
    result.first.fromExcel();
    return result;
  }

  inline std::pair<ExcelObj, int> tryCallExcel(int func, const ExcelObj& arg)
  {
    auto result = std::make_pair(ExcelObj(), 0);
    auto p = &arg;
    result.second = callExcelRaw(func, &result.first, 1, &p);
    result.first.fromExcel();
    return result;
  }
}


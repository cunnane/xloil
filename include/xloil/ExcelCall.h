#pragma once
#include <xlOil/ExcelObj.h>
#include <xlOil/Throw.h>
#include <vector>
#include <memory>
#include <list>

namespace xloil
{
  namespace detail
  {
    template <class TIterable> struct Splatter
    {
      Splatter(const TIterable& iter) : _obj(iter) {}
      const TIterable& operator()() const { return _obj; }
      const TIterable& _obj;
    };

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
      void add(detail::Splatter<TIter>&& splatter)
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

  /// <summary>
  /// Mimics pythons argument splat/unpack object for calls to
  /// <see cref="callExcel"/> by unpacking an iterable into function
  /// arguments.
  /// 
  /// <example>
  /// vector<double> vals;
  /// callExcel(xlfSum, unpack(vals));
  /// </example>
  /// </summary>
  template <class TIterable>
  auto unpack(const TIterable& iterable)
  {
    return detail::Splatter<TIterable>(iterable);
  }

  /// <summary>
  /// A wrapper around the Excel12 call. Avoid using directly unless for 
  /// performance reasons.
  /// </summary>
  XLOIL_EXPORT int callExcelRaw(
    int func, ExcelObj* result,
    size_t nArgs = 0,
    const ExcelObj** args = nullptr);

  /// <summary>
  /// 
  /// </summary>
  /// <typeparam name="T"></typeparam>
  /// <param name="func"></param>
  /// <param name="result"></param>
  /// <param name="nArgs"></param>
  /// <param name="firstArg"></param>
  /// <returns></returns>
  template <class T>
  inline int callExcelRaw(
    int func, ExcelObj* result,
    size_t nArgs,
    T firstArg)
  {
    const ExcelObj* pArgs[XL_MAX_UDF_ARGS];
    for (size_t i = 0; i < nArgs; ++i, ++firstArg)
      pArgs[i] = &*firstArg;
    return callExcelRaw(func, result, nArgs, pArgs);
  }

  /// <summary>
  /// Convenience wrapper for <see cref="callExcelRaw"/> for 
  /// a single argument
  /// </summary>
  inline int
    callExcelRaw(int func, ExcelObj* result, const ExcelObj* arg)
  {
    auto p = arg;
    return callExcelRaw(func, result, 1, &p);
  }

  /// <summary>
  /// Returns a reable error from the return code produced by
  /// <see cref="tryCallExcel"/>.
  /// </summary>
  XLOIL_EXPORT const wchar_t*
    xlRetCodeToString(int xlret, bool checkXllContext=true);

  /// <summary>
  /// If this error is thrown, Excel SDK documentation says you must
  /// immediately return control.
  /// </summary>
  class ExcelAbort : public std::runtime_error
  {
  public:
    ExcelAbort() : std::runtime_error("Excel abort") {}
  };

  /// <summary>
  /// Similar to <see cref="callExcel"/> but does not throw on failure.
  /// Rather returns a tuple (ExcelObj, int) where the second argument 
  /// is the return code <see cref="msxll::xlretSuccess"/>.
  /// </summary>
  template<typename... Args>
  inline std::pair<ExcelObj, int> 
    tryCallExcel(int func, Args&&... args) noexcept
  {
    auto result = std::make_pair(ExcelObj(), 0);
    detail::CallArgHolder holder(std::forward<Args>(args)...);
    result.second = callExcelRaw(func, &result.first, holder.nArgs(), holder.ptrToArgs());
    result.first.resultFromExcel();
    return result;
  }

  /// <summary>
  /// As for <see cref="tryCallExcel"/> but with no arguments.
  /// </summary>
  inline std::pair<ExcelObj, int> 
    tryCallExcel(int func)
  {
    auto result = std::make_pair(ExcelObj(), 0);
    result.second = callExcelRaw(func, &result.first);
    result.first.resultFromExcel();
    return result;
  }

  /// <summary>
  /// As for <see cref="tryCallExcel"/> but with a single ExcelObj argument.
  /// The separate implemenation gives some performance improvements.
  /// </summary>
  inline std::pair<ExcelObj, int> 
    tryCallExcel(int func, const ExcelObj& arg)
  {
    auto result = std::make_pair(ExcelObj(), 0);
    auto p = &arg;
    result.second = callExcelRaw(func, &result.first, 1, &p);
    result.first.resultFromExcel();
    return result;
  }

  /// <summary>
  /// Calls the specified Excel function number with the given arguments.
  /// Non-ExcelObj arguments are converted to ExcelObj types - this is 
  /// generally only possible for arithmetic and string types.
  /// 
  /// Throws an exeception if the call fails, otherwise returns the 
  /// result as an ExcelObj.
  /// </summary>
  template<typename... Args>
  inline ExcelObj callExcel(int func, Args&&... args)
  {
    auto [result, ret] = tryCallExcel(func, std::forward<Args>(args)...);
    switch (ret)
    {
    case msxll::xlretSuccess:
      break;
    case msxll::xlretAbort:
      throw ExcelAbort();
    default:
      XLO_THROW(L"Excel call failed: {0}", xlRetCodeToString(ret));
    }
    return result;
  }

  /// <summary>
  /// Convert an Excel built-in function name to a number for use <see cref="callExcel"/>
  /// </summary>
  /// <param name="name"></param>
  /// <returns>-1 if the name could not be found</returns>
  XLOIL_EXPORT int excelFuncNumber(const char* name);

  /// <summary>
  /// Convert an Excel built-in function number used in <see cref="callExcel"/> to its name string.
  /// </summary>
  /// <param name="number"></param>
  /// <returns>null if an invalid function number is passed</returns>
  XLOIL_EXPORT const char* excelFuncName(const unsigned number) noexcept;
}


#pragma once
#include <xloil/ExcelObj.h>
#include <xloil/Preprocessor.h>
#include <xloil/StaticRegister.h>
#include <xloil/ExcelArray.h>
#include <string>
#include <boost/preprocessor/repetition/enum.hpp>

namespace xloil
{
  namespace AutoBind
  {
    /// <summary>
    /// XLO_BIND attempts to register a wrapped C++ function. It can be called with 2, 3 or 4 
    /// arguments:
    /// 
    ///    * `XLO_BIND(FunctionName, NumArgs)`
    ///    * `XLO_BIND(FunctionName, NumArgs, Defaults)`
    ///    * `XLO_BIND(FunctionName, NumArgs, Defaults, Converters)`
    /// 
    /// The function is registered as `FunctionName` in Excel.
    /// 
    /// The `FunctionName` must be imported into the current namespace.
    /// 
    /// `Defaults` must be specified as `XLO_DEFAULTS(None, None, 1, 3.14)` where `None` indicates   
    /// no default.  The defaults do not have to match those specified by the C++ funding being bound.   
    /// 
    /// xlOil attempts to convert arguments from the `ExcelObj` type provided by Excel to the 
    /// type in the function definition by using the `ArgConvert` class.  The excel arguments must
    /// correspond 1-1 with the arguments in the bound function.
    ///  
    /// </summary>
#define XLO_BIND(...) XLO_EXPAND( _XLO_BIND_SELECT(__VA_ARGS__, _XLO_BIND4, _XLO_BIND3, _XLO_BIND2)(__VA_ARGS__) )

#define XLO_DEFAULTS(...) ::std::make_tuple(__VA_ARGS__)

    /// <summary>
    /// Can be used in the 4-arg version of XLO_BIND when no defaults are required
    /// </summary>
#define XLO_NODEFAULTS(NumArgs) ::std::make_tuple(BOOST_PP_ENUM(NumArgs, _XLO_NODEFAULTS_TEXT, None))

    /// <summary>
    /// Indicates an argument is not defaulted
    /// </summary>
    struct NoneType {};
    constexpr NoneType None;

    /// <summary>
    /// During auto-binding, arguments are convertered via `ArgConvert<T>(const ExcelObj& x)`, 
    /// so to add support for more argument types, provide specialisations of this class.
    /// A number of specialisations for standard C++ types are already given.
    /// </summary>
    template<class T> struct ArgConvert {};

    template<> struct ArgConvert<double>
    {
      double operator()(const ExcelObj& x) { return x.get<double>(); }
    };
    template<> struct ArgConvert<int>
    {
      int operator()(const ExcelObj& x) { return x.get<int>(); }
    };
    template<> struct ArgConvert<std::wstring>
    {
      std::wstring operator()(const ExcelObj& x) { return x.toString(); }
    };
    template<> struct ArgConvert<std::string>
    {
      std::string operator()(const ExcelObj& x) { return utf16ToUtf8(x.toString()); }
    };
    template<> struct ArgConvert<const wchar_t*>
    {
      const wchar_t* operator()(const ExcelObj& x) 
      { 
        _data = x.toString();
        return _data.c_str();
      }
      std::wstring _data;
    };

    template<> struct ArgConvert<std::vector<double>>
    {
      std::vector<double> operator()(const ExcelObj& obj) 
      { 
        ExcelArray array(obj);
        std::vector<double> result(array.size());
        std::transform(array.begin(), array.end(), result.begin(), [](auto& x)
          { return x.get<double>(); });
        return result;
      }
    };

    /// <summary>
    /// During auto-binding, the return value is convertered via `ReturnConvert<T>(T x)`, 
    /// so to add support for more return types, provide specialisations of this class.
    /// The call operator must return an un-owned `ExcelObj*`
    /// </summary>
    template<class T> struct ReturnConvert
    {
      ExcelObj* operator()(T val)
      {
        return new ExcelObj(val);
      }
    };

    template<> struct ReturnConvert<std::vector<double>>
    {
      ExcelObj* operator()(const std::vector<double>& val)
      {
        ExcelArrayBuilder builder(val.size(), 1);
        std::copy(val.begin(), val.end(), builder.begin());
        return new ExcelObj(builder.toExcelObj());
      }
    };

    namespace detail
    {
      template<template<typename> typename TArgConverter, typename T>
      struct Defaulting
      {
        auto operator()(const ExcelObj& x, T defaultVal)
        {
          if (!x.isMissing())
            return _converter(x);
          else
            return defaultVal;
        }

        auto operator()(const ExcelObj& x, NoneType)
        {
          return _converter(x);
        }

        TArgConverter<T> _converter;
      };

      template<template<typename> typename TArgConverter, template<typename> typename TReturnConverter, class TFunc>
      struct ConvertAndInvoke {};

      template <template<typename> typename TArgConverter, template<typename> typename TReturnConverter, typename ReturnType, typename... FuncArgs>
      struct ConvertAndInvoke<TArgConverter, TReturnConverter, ReturnType(FuncArgs...)>
      {
        template<typename F, typename Defaults, typename... ExcelArgs, std::size_t... I>
        auto impl(F func, Defaults defaults, std::index_sequence<I...>, ExcelArgs... args)
        {
          return TReturnConverter<ReturnType>()(
            std::invoke(func, 
              Defaulting<TArgConverter, std::remove_const_t<std::remove_reference_t<FuncArgs>>>()(
                args, std::get<I>(defaults))...
            ));
        }

        template<typename F, typename Defaults, typename... ExcelArgs, typename Indices = std::make_index_sequence<sizeof...(FuncArgs)>>
        auto operator()(F func, Defaults defaults, ExcelArgs... args)
        {
          return impl(func, defaults, Indices{}, args...);
        }
      };
    }

    struct DefaultConverters
    {
      template<class T> using Arg = ArgConvert<T>;
      template<class T> using Return = ReturnConvert<T>;
    };


#define _XLO_NODEFAULTS_TEXT(z, n, text) text

#define _XLO_BIND4(FUNC, NUM_ARGS, DEFAULTS, CONVERTERS) \
    XLO_FUNC_START(FUNC##_XLOIL(XLO_DECLARE_ARGS(NUM_ARGS, arg)))\
    { \
        auto defaults = DEFAULTS; \
        return ::xloil::AutoBind::detail::ConvertAndInvoke<CONVERTERS::Arg, CONVERTERS::Return, decltype(FUNC)>()( \
          FUNC, defaults, BOOST_PP_ENUM_SHIFTED_PARAMS(BOOST_PP_ADD(NUM_ARGS,1), arg)); \
    } XLO_FUNC_END(FUNC##_XLOIL).name(XLO_WSTR(FUNC))

#define _XLO_BIND3(FUNC, NUM_ARGS, DEFAULTS) _XLO_BIND4(FUNC, NUM_ARGS, DEFAULTS, ::xloil::AutoBind::DefaultConverters)

#define _XLO_BIND2(FUNC, NUM_ARGS) _XLO_BIND3(FUNC, NUM_ARGS, XLO_NODEFAULTS(NUM_ARGS))

#define _XLO_BIND_SELECT(_1,_2,_3,_4, NAME,...) NAME

  }
}
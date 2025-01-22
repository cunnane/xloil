#pragma once
#include <boost/preprocessor/repeat_from_to.hpp>
#include <boost/preprocessor/arithmetic.hpp>
#include <boost/preprocessor/repetition/enum_shifted_params.hpp>
#include <boost/preprocessor/tuple/elem.hpp>
#include <boost/preprocessor/cat.hpp>
#include <boost/preprocessor/stringize.hpp>
#include <xloil/Version.h>

namespace xloil { class ExcelObj; }

/// <summary>  
/// https://stackoverflow.com/questions/5134523
/// </summary>
#define XLO_EXPAND( x ) x

/// <summary>  
/// Stringifies the given argument to a wide string literal
/// </summary>
#define XLO_WSTR(s) BOOST_PP_CAT(L,BOOST_PP_CAT(BOOST_PP_EMPTY(),BOOST_PP_STRINGIZE(s)))

/// <summary>
/// Stringifies the given argument to a narrow string literal
/// </summary>
#define XLO_STR(s) BOOST_PP_STRINGIZE(s)

/// <summary>
/// Use in a function declaration to declare <code>const ExcelObj&amp;</code> arguments 
/// named prefix1, prefix2, ..., prefixN
/// </summary>
#define XLO_DECLARE_ARGS(N, prefix) BOOST_PP_ENUM_SHIFTED_PARAMS(BOOST_PP_ADD(N,1), const ::xloil::ExcelObj& prefix)

/// <summary>
/// Returns a comma-separated list of argument addresses: &amp;prefix1, &amp;prefix2, ..., &amp;prefixN.
/// Useful to create an array of function arguments.
/// </summary>
#define XLO_ARG_PTRS(N, prefix) BOOST_PP_ENUM_SHIFTED_PARAMS(BOOST_PP_ADD(N,1), &prefix)

/// <summary>
/// Returns a list of argument values and their names for use with <see cref="xloil::ProcessArgs"/>.
/// </summary>
#define XLO_ARGS_LIST(N, prefix) BOOST_PP_REPEAT_FROM_TO(1, BOOST_PP_ADD(1, N), XLO_ARGS_LIST_IMPL, prefix)
#define XLO_ARGS_LIST_IMPL(z, N, prefix) BOOST_PP_COMMA_IF(BOOST_PP_SUB(N, 1)) prefix##N, XLO_WSTR(BOOST_PP_CAT(prefix,N))

/// <summary>
/// Writes repeated arg descriptors in the form `.arg(prefixN, help)`
/// </summary>
#define XLO_WRITE_ARG_HELP(N, prefix, help) \
  BOOST_PP_REPEAT_FROM_TO(1, BOOST_PP_ADD(1, N), XLO_WRITE_ARG_HELP_I, (prefix, help))
#define XLO_WRITE_ARG_HELP_I(z, N, data) .arg(XLO_WSTR(BOOST_PP_CAT(BOOST_PP_TUPLE_ELEM(0, data),N)), BOOST_PP_TUPLE_ELEM(1, data))

/// <summary>
/// As <see cref="XLO_ARG_PTRS"/>, but runs <see cref="xloil::objectCacheExpand"/> on each argument.
/// Useful to create an array of function arguments.
/// </summary>
#define XLO_CACHE_ARG_PTRS_I(z, N, data) &::xloil::objectCacheExpand(BOOST_PP_CAT(data ## N)) BOOST_PP_COMMA_IF(N)
#define XLO_CACHE_ARG_PTRS(N, prefix) BOOST_PP_REPEAT_FROM_TO(1, BOOST_PP_ADD(1, N), XLO_CACHE_ARG_PTRS_I, prefix)

#define _XLOIL_VERSION_STR (XLOIL_MAJOR_VERSION)(.)(XLOIL_MINOR_VERSION)(.)(XLOIL_PATCH_VERSION)
#define XLOIL_VERSION_STR XLO_WSTR(BOOST_PP_SEQ_CAT(OUR_VERSION))

namespace xloil
{
  class ExcelObj;

  /// <summary>
  /// Iterates over a number of ExcelObj arguments, applying a function to each.
  /// Best illustrated with an example:
  /// 
  /// <example>
  ///   ProcessArgs([&str](auto iArg, auto argVal, auto argName)
  ///   {
  ///      str += wstring(argName) + ": " + argVal.toString() + "\n";
  ///   }, XLO_ARGS_LIST(8, arg));
  /// </example>
  /// 
  /// ProcessArgs will accept lambdas which do not contain the <code>Arg</code>
  /// or <code>argName</code> arguments.
  ///  
  /// </summary>
  template<int N = 0, class TFunc, class... Args>
  auto ProcessArgs(TFunc func, const ExcelObj& argVal, const wchar_t* argName, Args&&...args)
    -> decltype(func(N, argVal, argName))
  {
    func(N, argVal, argName);
    ProcessArgs<N + 1>(func, args...);
  }
  template<int N, class TFunc>
  auto ProcessArgs(TFunc func, const ExcelObj& argVal, const wchar_t* argName)
    -> decltype(func(N, argVal, argName))
  {
    func(N, argVal, argName);
  }

  template<class TFunc, class... Args>
  auto ProcessArgs(TFunc func, const ExcelObj& argVal, const wchar_t* argName, Args&&... args)
    -> decltype(func(argVal, argName))
  {
    func(argVal, argName);
    ProcessArgs(func, args...);
  }

  template<class TFunc>
  auto ProcessArgs(TFunc func, const ExcelObj& argVal, const wchar_t* argName)
    -> decltype(func(argVal, argName))
  {
    func(argVal, argName);
  }

  template<class TFunc, class... Args>
  auto ProcessArgs(TFunc func, const ExcelObj& argVal, const wchar_t* /*argName*/, Args&&... args)
    -> decltype(func(argVal))
  {
    func(argVal);
    ProcessArgs(func, args...);
  }

  template<class TFunc>
  auto ProcessArgs(TFunc func, const ExcelObj& argVal, const wchar_t* /*argName*/)
    -> decltype(func(argVal))
  {
    func(argVal);
  }
}

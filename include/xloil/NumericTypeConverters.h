#pragma once
#include <xlOil/TypeConverters.h>
#include <xlOil/StringUtils.h>
#include <xlOil/Throw.h>

namespace xloil
{
  namespace conv
  {
    template<class TResult = double>
    struct ToDouble : public FromExcelBase<TResult>
    {
      using result_t = TResult;
      using FromExcelBase<TResult>::operator();
      result_t operator()(int x) const { return double(x); }
      result_t operator()(bool x) const { return double(x); }
      result_t operator()(double x) const { return x; }
      result_t operator()(CellError err) const
      {
        using namespace msxll;
        if (0 != ((int)err & (xlerrNull | xlerrDiv0 | xlerrNum | xlerrNA)))
          return std::numeric_limits<double>::quiet_NaN();
        XLO_THROW("Could not convert error to double");
      }
    };
    template<class TResult = int>
    struct ToInt : public FromExcelBase<TResult>
    {
      using result_t = TResult;
      using FromExcelBase<TResult>::operator();

      result_t operator()(int x) const { return x; }
      result_t operator()(bool x) const { return int(x); }
      result_t operator()(double x) const
      {
        int i;
        if (floatingToInt(x, i))
          return i;
        XLO_THROW("Could not convert: number not an exact integer");
      }
    };

    /// Converts to bool using Excel's standard coercions for numeric types (x != 0)
    template<class TResult = bool>
    struct ToBool : public FromExcelBase<TResult>
    {
      using result_t = TResult;
      using FromExcelBase<TResult>::operator();
      result_t operator()(int x) const { return x != 0; }
      result_t operator()(bool x) const { return x; }
      result_t operator()(double x) const { return x != 0.0; }
    };
  }
  /// <summary>
  /// Implementation of FromExcel which converts an ExcelObj to a double
  /// or throws if this is not possible
  /// </summary>
  inline auto ToDouble()           { return std::move(FromExcel<conv::ToDouble<>>()); }
  inline auto ToDouble(double def) { return std::move(FromExcelDefaulted<conv::ToDouble<>>(def)); }
  
  /// <summary>
  /// Implementation of FromExcel which converts an ExcelObj to an int
  /// or throws if this is not possible
  /// </summary>
  inline auto ToInt()        { return std::move(FromExcel<conv::ToInt<>>()); }
  inline auto ToInt(int def) { return std::move(FromExcelDefaulted<conv::ToInt<>>(def)); }
  
  /// <summary>
  /// Implementation of FromExcel which converts an ExcelObj to a bool
  /// or throws if this is not possible
  /// </summary>
  inline auto ToBool()         { return std::move(FromExcel<conv::ToBool<>>()); }
  inline auto ToBool(bool def) { return std::move(FromExcelDefaulted<conv::ToBool<>>(def)); }
}
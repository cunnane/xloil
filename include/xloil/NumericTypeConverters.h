#pragma once
#include "TypeConverters.h"
#include <xlOil/StringUtils.h>
#include <xlOil/Throw.h>

namespace xloil
{
  namespace conv
  {
    template<class TResult = double>
    struct ToDouble : public FromExcelBase<TResult>
    {
      using result = TResult;
      using FromExcelBase::operator();
      result operator()(int x) const { return double(x); }
      result operator()(bool x) const { return double(x); }
      result operator()(double x) const { return x; }
      result operator()(CellError err) const
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
      using result = TResult;
      using FromExcelBase::operator();

      result operator()(int x) const { return x; }
      result operator()(bool x) const { return int(x); }
      result operator()(double x) const
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
      using result = TResult;
      using FromExcelBase::operator();
      result operator()(int x) const { return x != 0; }
      result operator()(bool x) const { return x; }
      result operator()(double x) const { return x != 0.0; }
    };
  }
  /// <summary>
  /// Implementation of FromExcel which covnverts an ExcelObj to a double
  /// or throws if this is not possible
  /// </summary>
  using ToDouble = FromExcel<conv::ToDouble<double>>;
  /// <summary>
  /// Implementation of FromExcel which covnverts an ExcelObj to an int
  /// or throws if this is not possible
  /// </summary>
  using ToInt = FromExcel<conv::ToInt<int>>;
  /// <summary>
  /// Implementation of FromExcel which covnverts an ExcelObj to a bool
  /// or throws if this is not possible
  /// </summary>
  using ToBool = FromExcel<conv::ToBool<bool>>;
}
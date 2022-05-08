#pragma once
#include <xlOil/TypeConverters.h>
#include <xlOil/StringUtils.h>
#include <xlOil/Throw.h>

namespace xloil
{
  namespace conv
  {
    template<class TValue, class TResult> struct ExcelValToType {};

    template<class TResult>
    struct ExcelValToType<double, TResult> : public ExcelValVisitor<TResult>
    {
      using result_t = TResult;
      using base = ExcelValVisitor<TResult>;
      using base::operator();

      result_t operator()(int x) const { return double(x); }
      result_t operator()(bool x) const { return double(x); }
      result_t operator()(double x) const { return x; }
      result_t operator()(CellError err) const
      {
        using namespace msxll;
        if (0 != ((int)err & (xlerrNull | xlerrDiv0 | xlerrNum | xlerrNA)))
          return std::numeric_limits<double>::quiet_NaN();
        return base::operator()(err);
      }
    };

    template<class TResult>
    struct ExcelValToType<int, TResult> : public ExcelValVisitor<TResult>
    {
      using result_t = TResult;
      using base = ExcelValVisitor<TResult>;
      using base::operator();

      result_t operator()(int x) const { return x; }
      result_t operator()(bool x) const { return int(x); }
      result_t operator()(double x) const
      {
        int i;
        if (floatingToInt(x, i))
          return i;
        return base::operator()(x);
      }
    };

    /// Converts to bool using Excel's standard coercions for numeric types (x != 0)
    template<class TResult>
    struct ExcelValToType<bool, TResult> : public ExcelValVisitor<TResult>
    {
      using result_t = TResult;
      using ExcelValVisitor<TResult>::operator();
      result_t operator()(int x) const { return x != 0; }
      result_t operator()(bool x) const { return x; }
      result_t operator()(double x) const { return x != 0.0; }
    };

    template<>
    struct ExcelValToType<std::wstring_view, std::wstring_view>
      : public ExcelValVisitor<std::wstring_view>
    {
      using result_t = std::wstring_view;
      using ExcelValVisitor<std::wstring_view>::operator();
      result_t operator()(PStringRef str) const { return str; }
    };

    namespace detail
    {
      // Need this indirection layer as we can't partially specialise
      // an alias, i.e. a `using` declaration.
      template<class T>
      struct ExcelValToTypeSpecialisation { using type = ExcelValToType<T, T>; };
      template<class T>
      struct ExcelValToTypeSpecialisation<std::optional<T>> { using type = ExcelValToType<T, std::optional<T>>; };
    }

    template<class T> using ToType = 
      typename detail::ExcelValToTypeSpecialisation<T>::type;

  }
}
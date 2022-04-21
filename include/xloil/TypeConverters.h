#pragma once
#include <xlOil/ExcelObj.h>
#include <xlOil/Throw.h>

namespace xloil { class ExcelRef; }
namespace xloil
{
  /// <summary>
  /// Interface class for converters which take an ExcelObj and output a
  /// language-specific type
  /// </summary>
  template <class TResult>
  class IConvertFromExcel
  {
  public:
    using result_type = TResult;
    using const_result_ptr = const typename std::remove_pointer<TResult>::type*;
    virtual ~IConvertFromExcel() {}
    virtual result_type operator()(
      const ExcelObj& xl, 
      const_result_ptr defaultVal = nullptr) const = 0;
  };

  /// <summary>
  /// Interface class for converters which take a language-specific type
  /// and produce an ExcelObj
  /// </summary>
  template <class TSource>
  class IConvertToExcel
  {
  public:
    virtual ~IConvertToExcel() {}
    virtual ExcelObj operator()(const TSource& obj) const = 0;
  };

  /// <summary>
  /// Holder class which allows the type converter implementation to select
  /// objects which represent arrays
  /// </summary>
  struct ArrayVal
  {
    const ExcelObj& obj;
    operator const ExcelObj& () const { return obj; }
  };

  /// <summary>
  /// Holder class which allows the type converter implementation to select
  /// objects which represent range references
  /// </summary>
  struct RefVal
  {
    const ExcelObj& obj;
    operator const ExcelObj& () const { return obj; }
  };

  /// <summary>
  /// Indicates a missing value to a type converter implementation
  /// </summary>
  struct MissingVal
  {};

  /// <summary>
  /// Handles the switch on the type of the ExcelObj and dispatches to an
  /// overload of the functor's operator(). Called via <see cref="FromExcel"/>.
  /// </summary>
  template<class TFunc>
  auto visitExcelObj(
    const ExcelObj& xl, 
    TFunc functor)
  {
    try
    {
      switch (xl.type())
      {
      case ExcelType::Int:     return functor(xl.val.w);
      case ExcelType::Bool:    return functor(xl.val.xbool != 0);
      case ExcelType::Num:     return functor(xl.val.num);
      case ExcelType::Str:     return functor(xl.asPString());
      case ExcelType::Multi:   return functor(ArrayVal{ xl });
      case ExcelType::Missing: return functor(MissingVal());
      case ExcelType::Err:     return functor(CellError(xl.val.err));
      case ExcelType::Nil:     return functor(nullptr);
      case ExcelType::SRef:
      case ExcelType::Ref:
        return functor(RefVal{ xl });
      default:
        XLO_THROW("Unexpected XL type");
      }
    }
    catch (const std::exception& e)
    {
      XLO_THROW(L"Failed reading {0}: {1}", 
        xl.toStringRepresentation(), 
        utf8ToUtf16(e.what()));
    }
  }

  /// <summary>
  /// Provides the default implementation (which is generally an error)
  /// for conversion functors to be used in <see cref="FromExcel"/> or 
  /// <see cref="visitExcelObj"/>
  /// </summary>
  template <class TResult>
  class FromExcelBase
  {
  public:
    template <class T>
    TResult operator()(T) const 
    { 
      throw std::runtime_error("Cannot convert to required type");
    }
  };

  template <class TResult>
  class FromExcelBase<std::optional<TResult>>
  {
  public:
    template <class T>
    auto operator()(T) const
    {
      return std::optional<TResult>();
    }
  };

  template <class TResult>
  class FromExcelBase<TResult*>
  {
  public:
    template <class T>
    TResult* operator()(T) const
    {
      return nullptr;
    }
  };

  template<class TBase>
  class FromExcelHandleMissing : public TBase
  {
    const typename TBase::result_t& _defaultValue;

  public:
    template <class...Args>
    FromExcelHandleMissing(const typename TBase::result_t& defaultVal, Args...args)
      : TBase(std::forward<Args>(args)...)
      , _defaultValue(defaultVal)
    {}

    using TBase::operator();
    auto operator()(MissingVal) const
    {
      return _defaultValue;
    }
  };

  /// <summary>
  /// Creates a functor which applies a type conversion implementation 
  /// functor to an ExcelObj
  /// </summary>
  template<class TImpl>
  class FromExcel
  {
    TImpl _impl;
  public:

    template <class...Args>
    FromExcel(Args...args)
      : _impl(std::forward<Args>(args)...)
    {}

    using return_type = decltype(_impl(nullptr));

    /// <summary>
    /// Applies the type conversion implementation functor to the ExcelObj.
    /// If provided, the default value is returned if the ExcelObj is of
    /// type Missing.
    /// </summary>
    return_type operator()(const ExcelObj& xl) const
    {
      return visitExcelObj(xl, _impl);
    }
  };

  template<class TImpl>
  using FromExcelDefaulted = FromExcel<FromExcelHandleMissing<TImpl>>;
}
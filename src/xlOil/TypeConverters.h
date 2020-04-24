#pragma once
#include "ExcelObj.h"
#include <xlOil/Throw.h>

namespace xloil { class ExcelRange; }
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
    virtual result_type operator()(const ExcelObj& xl, const_result_ptr defaultVal = nullptr) const = 0;
  };

  /// <summary>
  /// Interface class for converters which take a language-specific type
  /// and produce an ExcelObj
  /// </summary>
  template <class TSource>
  class IConvertToExcel
  {
  public:
    virtual ExcelObj operator()(const TSource& obj) const = 0;
  };

  /// <summary>
  /// Wrapper around an Excel->Language type converter which handles the 
  /// switch on the type of the ExcelObj.
  /// </summary>
  template <class TResult, class TImpl>
  struct FromExcelDispatcher
  {
    using result_type = TResult;

    const TImpl& _impl() const { return static_cast<const TImpl&>(*this); }

    using const_return_ptr = const typename std::remove_pointer<TResult>::type*;

    TResult operator()(const ExcelObj& xl, const_return_ptr defaultVal = nullptr) const
    {
      switch (xl.type())
      {
      case ExcelType::Int: return _impl().fromInt(xl.val.w);
      case ExcelType::Bool: return _impl().fromBool(xl.val.xbool != 0);
      case ExcelType::Num: return _impl().fromDouble(xl.val.num);
      case ExcelType::Str: return _impl().fromString(xl.val.str + 1, xl.val.str[0]);
      case ExcelType::Multi: return _impl().fromArray(xl);
      case ExcelType::Missing: return _impl().fromMissing(defaultVal);
      case ExcelType::Err: return _impl().fromError(CellError(xl.val.err));
      case ExcelType::Nil: return _impl().fromEmpty(defaultVal);
      case ExcelType::SRef:
      case ExcelType::Ref:
        return _impl().fromRef(xl);
      default:
        XLO_THROW("Unexpected XL type");
      }
    }
  };

  template<class Super, class This>
  using NotNull = typename std::conditional<std::is_same<Super, nullptr_t>::value, This, Super>::type;

  /// <summary>
  /// Does the actual work of conversion from Excel to a language 
  /// type.  The fromXXX methods should be overriden for as many
  /// ExcelObj data types as make sense for the conversion being performed.
  /// </summary>
  template <class TResult, class TSuper=nullptr_t>
  class FromExcelBase 
    : public FromExcelDispatcher<TResult, NotNull<TSuper, FromExcelBase<TResult, nullptr_t>>>
  {
  public:
    TResult fromInt(int x) const { return error(); }
    TResult fromBool(bool x) const { return error(); }
    TResult fromDouble(double x) const { return error(); }
    TResult fromArray(const ExcelObj&) const { return error(); }
    TResult fromArrayObj(const ExcelArray&) const { return error(); }
    TResult fromString(const wchar_t* /*buf*/, size_t /*len*/) const { return error(); }
    TResult fromError(CellError) const { return error(); }
    TResult fromEmpty(const TResult* /*defaultVal*/) const { return error(); }
    TResult fromMissing(const TResult* defaultVal) const 
    { 
      if (defaultVal)
        return *defaultVal;
      XLO_THROW("Missing argument");
    }
    TResult fromRef(const ExcelObj&) const { return error(); }
    TResult fromRef(const ExcelRange&) const { return error(); }

    TResult error() const { XLO_THROW("Cannot convert"); }
  };

  /// <summary>
  /// Specialised ConverterImpl which returns nullptr instead of throwing an
  /// error when a type conversion cannot be performed, i.e. the fromXXX 
  /// function has not be overriden for the supplied ExcelObj type
  /// </summary>
  template <class TResult, class TSuper>
  class FromExcelBase<TResult*, TSuper>
    : public FromExcelDispatcher<TResult*, NotNull<TSuper, FromExcelBase<TResult*, nullptr_t>>>
  {
  public:
    TResult* fromInt(int) const { return nullptr; }
    TResult* fromBool(bool) const { return nullptr; }
    TResult* fromDouble(double) const { return nullptr; }
    // Need to give this a different name or it seems to break C++ overload 
    // resolution. Unless the rules change for some reason in templates.
    TResult* fromArrayObj(const ExcelArray&) const { return nullptr; }
    TResult* fromArray(const ExcelObj&) const { return nullptr; }
    TResult* fromString(const wchar_t* /*buf*/, size_t /*len*/) const { return nullptr; }
    TResult* fromError(CellError) const { return nullptr; }
    TResult* fromEmpty(const TResult* defaultVal) const { return const_cast<TResult*>(defaultVal); }
    TResult* fromMissing(const TResult* defaultVal) const 
    { 
      if (defaultVal)
        return const_cast<TResult*>(defaultVal);
      XLO_THROW("Missing argument");
    }
    TResult* fromRef(const ExcelObj& obj) const { return nullptr; }
    TResult* fromRefObj(const ExcelRange& rng) const { return nullptr; }
  };


  template <class TResult, class TSuper=nullptr_t>
  struct CacheConverter : public FromExcelBase<TResult, NotNull<TSuper, CacheConverter<TResult, nullptr_t>>>
  {
    auto fromString(const wchar_t* buf, size_t len) const
    {
      if (maybeObjectCacheReference(buf, len))
      {
        std::shared_ptr<const ExcelObj> obj;
        if (xloil::fetchCacheObject(buf, len, obj))
          return _impl()(*obj);
      }
      return FromExcelBase::fromString(buf, len);
    }
  };
}
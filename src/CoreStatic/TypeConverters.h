#pragma once
#include "ExcelObj.h"
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
    using return_type = TResult;
    using const_return_ptr = const typename std::remove_pointer<TResult>::type*;
    virtual return_type operator()(const ExcelObj& xl, const_return_ptr defaultVal = nullptr) const = 0;
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
  /// Implementation of IConvertFromExcel which just wraps an Impl
  /// class to do the work. Impl classes are non virtual, this class
  /// exists to bridge to virtual overide operator().
  /// </summary>
  template <class TImpl>
  class ConvertFromExcel : public IConvertFromExcel<typename TImpl::return_type>
  {
  private:
    TImpl _impl;

  public:
    using return_type = typename IConvertFromExcel::return_type;
    using const_return_ptr = typename IConvertFromExcel::const_return_ptr;

    template <class...Args>
    ConvertFromExcel(Args&&...args)
      : _impl(std::forward<Args>(args)...) {}

    virtual return_type operator()(const ExcelObj& xl, const_return_ptr defaultVal = nullptr) const override
    {
      return _impl(xl, defaultVal);
    }
  };

  /// <summary>
  /// Wrapper around an Excel->Language type converter which handles the 
  /// switch on the type of the ExcelObj.
  /// </summary>
  template <class TImpl>
  class FromExcel
  {
    TImpl _impl;
  public:
    using return_type = decltype(_impl.fromInt(1)); // Choice of fromInt is arbitrary
    using const_return_ptr = const typename std::remove_pointer<return_type>::type*;

    template <class...Args>
    FromExcel(Args&&...args)
      : _impl(std::forward<Args>(args)...) {}

    return_type operator()(const ExcelObj& xl, const_return_ptr defaultVal = nullptr) const
    {
      switch (xl.type())
      {
      case ExcelType::Int: return _impl.fromInt(xl.val.w);
      case ExcelType::Bool: return _impl.fromBool(xl.val.xbool != 0);
      case ExcelType::Num: return _impl.fromDouble(xl.val.num);
      case ExcelType::Str: return _impl.fromString(xl.val.str + 1, xl.val.str[0]);
      case ExcelType::Multi: return _impl.fromArray(xl);
      case ExcelType::Missing: return _impl.fromMissing(defaultVal);
      case ExcelType::Err: return _impl.fromError(CellError(xl.val.err));
      case ExcelType::Nil: return _impl.fromEmpty(defaultVal);
      default:
        XLO_THROW("Unexpected XL type");
      }
    }
  };

  /// <summary>
  /// Does the actual work of conversion from Excel to a language 
  /// type.  The fromXXX methods should be overriden for as many
  /// ExcelObj data types as make sense for the conversion being performed.
  /// </summary>
  template <class TResult>
  class ConverterImpl
  {
  public:
    TResult fromInt(int x) const { return error(); }
    TResult fromBool(bool x) const { return error(); }
    TResult fromDouble(double x) const { return error(); }
    TResult fromArray(const ExcelObj& obj) const { return error(); }
    TResult fromString(const wchar_t* buf, size_t len) const { return error(); }
    TResult fromError(CellError err) const { return error(); }
    TResult fromEmpty(const TResult* defaultVal) const { return error(); }
    TResult fromMissing(const TResult* defaultVal) const 
    { 
      if (defaultVal)
        return *defaultVal;
      XLO_THROW("Missing argument");
    }

    TResult error() const { XLO_THROW("Cannot convert"); }
  };

  /// <summary>
  /// Specialised ConverterImpl which returns nullptr instead of throwing an
  /// error when a type conversion cannot be performed, i.e. the fromXXX 
  /// function has not be overriden for the supplied ExcelObj type
  /// </summary>
  template <class TResult>
  class ConverterImpl<TResult*>
  {
  public:
    TResult* fromInt(int x) const { return nullptr; }
    TResult* fromBool(bool x) const { return nullptr; }
    TResult* fromDouble(double x) const { return nullptr; }
    TResult* fromArray(const ExcelObj& obj) const { return nullptr; }
    TResult* fromString(const wchar_t* buf, size_t len) const { return nullptr; }
    TResult* fromError(CellError err) const { return nullptr; }
    TResult* fromEmpty(const TResult* defaultVal) const { return const_cast<TResult*>(defaultVal); }
    TResult* fromMissing(const TResult* defaultVal) const 
    { 
      if (defaultVal)
        return const_cast<TResult*>(defaultVal);
      XLO_THROW("Missing argument");
    }
  };
}
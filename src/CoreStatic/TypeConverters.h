#pragma once
#include "ExcelObj.h"
namespace xloil
{
  template <class TResult>
  class IConvertFromExcel
  {
  public:
    using return_type = TResult;
    using const_return_ptr = const typename std::remove_pointer<TResult>::type*;
    virtual return_type operator()(const ExcelObj& xl, const_return_ptr defaultVal = nullptr) const = 0;
  };

  template <class TSource>
  class IConvertToExcel
  {
  public:
    virtual ExcelObj operator()(const TSource& obj) const = 0;
  };

  template <class TImpl>
  class FromExcel
  {
    TImpl _impl;
  public:
    using return_type = decltype(_impl.fromInt(1));
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

  // TODO: should default fromString do a cache lookup?
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

  template <class TImpl>
  class ConvertFromExcel : public IConvertFromExcel<typename TImpl::return_type>
  {
    TImpl _impl;

  public:
    using base_type = ConvertFromExcel;
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
}
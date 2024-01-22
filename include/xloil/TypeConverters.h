#pragma once

#ifndef XLOIL_XLOPER
#error "Don't include this directly, include ExcelObj.h"
#endif

#include <optional>
#include <stdexcept>

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
      const_result_ptr defaultVal = nullptr) = 0;
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

  template <class TSource>
  class IConvertToExcel<TSource*>
  {
  public:
    virtual ~IConvertToExcel() {}
    virtual ExcelObj operator()(const TSource* obj) const = 0;
  };

  /// <summary>
  /// Provides the default implementation (which is generally an error)
  /// for conversion functors to be used in <see cref="FromExcel"/> or 
  /// <see cref="visitExcelObj"/>
  /// </summary>
  template <class TResult>
  class ExcelValVisitor
  {
  public:
    using return_type = TResult;
    template <class T> return_type operator()(T) const
    { 
      throw std::runtime_error("Cannot convert to required type");
    }
  };

  template <class TResult>
  class ExcelValVisitor<std::optional<TResult>>
  {
  public:
    using return_type = std::optional<TResult>;
    constexpr return_type operator()(...) const
    {
      return std::optional<TResult>();
    }
  };

  template <class TResult>
  class ExcelValVisitor<TResult*>
  {
  public:
    using return_type = TResult*;
    constexpr return_type operator()(...) const
    {
      return nullptr;
    }
  };

  template<class TVisitor>
  class ExcelValVisitorDefaulted 
    : public ExcelValVisitor<typename TVisitor::return_type>
  {
    using base = ExcelValVisitor<typename TVisitor::return_type>;
    using default_type = const typename TVisitor::return_type&;

    TVisitor _visitor;
    default_type _defaultValue;

  public:
    template <class...Args>
    ExcelValVisitorDefaulted(
      TVisitor visitor,
      default_type defaultVal)
      : _visitor(visitor)
      , _defaultValue(defaultVal)
    {}

    template<class T> auto operator()(T x) const
    {
      return _visitor(std::forward<T>(x));
    }

    auto operator()(MissingVal) const
    {
      return _defaultValue;
    }
  };
}
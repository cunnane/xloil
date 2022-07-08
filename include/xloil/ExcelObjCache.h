#pragma once
#include <xloil/ExcelObj.h>
#include <xloil/ObjectCache.h>
#include <memory>

namespace xloil
{
  /// <summary>
  /// The CacheUniquifier character for the Excel Object Cache
  /// </summary>
  template<>
  struct CacheUniquifier<std::unique_ptr<const ExcelObj>>
  {
    static constexpr wchar_t value = L'\x6C38';
    static constexpr auto tag = L"XLOBJ ";
  };

  template struct XLOIL_EXPORT ObjectCacheFactory<std::unique_ptr<const ExcelObj>>;

  /// <summary>
  /// If the argument is a string referencing an <see cref="ExcelObj"/>
  /// in the cache, the cached object is returned, otherwise the argument
  /// object is returned.
  /// </summary>
  /// <param name="obj"></param>
  /// <returns></returns>
  inline const ExcelObj& cacheCheck(const ExcelObj& obj)
  {
    if (obj.isType(ExcelType::Str))
    {
      auto cacheVal = getCached<ExcelObj>(obj.cast<PStringRef>().view());
      if (cacheVal)
        return *cacheVal;
    }
    return obj;
  }

  /// <summary>
  /// Runs <see cref="xloil::cacheCheck"/> on each supplied argument in-place. 
  /// Usage:
  /// <code>
  ///    ExcelObj *arg1, *arg2;
  ///    cacheCheck(arg1, arg2);
  ///    ExcelArray array1(arg1);
  ///    ...
  /// </code>
  /// </summary>
  /// <typeparam name="...Args"></typeparam>
  /// <param name="first"></param>
  /// <param name="...more"></param>
  template<class...Args>
  inline void cacheCheck(const ExcelObj*& first, Args&&... more)
  {
    first = &cacheCheck(*first);
    cacheCheck(std::forward<Args>(more)...);
  }

  inline void cacheCheck(const ExcelObj*& obj)
  {
    obj = &cacheCheck(*obj);
  }

  /// <summary>
  /// Function object which runs <see cref="xloil::cacheCheck"/>
  /// </summary>
  struct CacheCheck
  {
    void operator()(const ExcelObj*& obj) { cacheCheck(obj); }
    template<class...Args>
    void operator()(Args&&... args) { cacheCheck(std::forward<Args>(args)...); }
  };

  /// <summary>
  /// Wraps a type conversion functor, interepting the string conversion to
  /// look for a cache reference.  If found, calls the wrapped functor on the
  /// cache object, otherwise passes the string through.
  /// </summary>
  template<class TBase>
  struct CacheConverter : public TBase
  {
    template <class...Args>
    CacheConverter(Args&&...args)
      : TBase(std::forward<Args>(args)...)
    {}

    using TBase::operator();
    auto operator()(const PStringRef& str) const
    {
      const ExcelObj* obj = getCached<ExcelObj>(str.view());
      if (obj)
        return obj->visit((TBase&)(*this));
      
      return TBase::operator()(str);
    }
  };
}
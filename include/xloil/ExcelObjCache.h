#pragma once
#include <xloil/ExcelObj.h>
#include <xloil/ObjectCache.h>
#include <memory>

namespace xloil
{
  template<>
  struct CacheUniquifier<std::unique_ptr<const ExcelObj>>
  {
    static constexpr wchar_t value = L'\x6C38';
  };

  template struct XLOIL_EXPORT ObjectCacheFactory<std::unique_ptr<const ExcelObj>>;

  using objectCacheUnquifier = CacheUniquifier<std::unique_ptr<const ExcelObj>>;
 
  /// <summary>
  /// Adds an ExcelObj to the object cache and returns an reference string
  /// (as an ExcelObj) based on the currently executing cell
  /// </summary>


  inline const ExcelObj& objectCacheExpand(const ExcelObj& obj)
  {
    if (obj.isType(ExcelType::Str))
    {
      auto cacheVal = getCached<ExcelObj>(obj.asPString().view());
      if (cacheVal)
        return *cacheVal;
    }
    return obj;
  }


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
    auto operator()(const PStringView<>& str) const
    {
      const ExcelObj* obj = getCached<ExcelObj>(str.view());
      if (obj)
        return visitExcelObj(*obj, (TBase&)(*this));
      
      return TBase::operator()(str);
    }
  };

}
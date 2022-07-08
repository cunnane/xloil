#pragma once
#include <xloil/ExcelObj.h>

typedef struct tagVARIANT VARIANT;
typedef struct tagSAFEARRAY SAFEARRAY;

namespace xloil
{
  namespace COM
  {
    namespace detail
    {
      class SafeArrayAccessorBase
      {
      public:
        SafeArrayAccessorBase(SAFEARRAY* pArr);
        ~SafeArrayAccessorBase();

        const size_t dimensions, rows, cols;

      protected:
        SAFEARRAY* _ptr;
        void* _data;
      };
    }
    /// <summary>
    /// Fast access to a SafeArray's internal data store. The template 
    /// parameter must match the data type of the SafeArray (this is not
    /// checked) or core will be dumped.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    template<class T>
    class SafeArrayAccessor : public detail::SafeArrayAccessorBase
    {
    public:
      using SafeArrayAccessorBase::SafeArrayAccessorBase;

      T& operator()(const size_t i, const size_t j)
      {
        return data()[j * rows + i];
      }

      const T& operator()(const size_t i, const size_t j) const
      {
        return data()[j * rows + i];
      }

      /// <summary>
      /// Returns the size of the sub-array trimmed to the last non-empty 
      /// (not Nil, \#N/A or "") row and column.
      /// </summary>
      std::pair<size_t, size_t> trimmedSize() const
      {
        return std::pair(rows, cols);
      }

    private:
      T* data() const { return (T*)_data; }
    };

    template<> std::pair<size_t, size_t> SafeArrayAccessor<VARIANT>::trimmedSize() const;

    void excelObjToVariant(VARIANT* v, const ExcelObj& obj, bool allowRange = false);
    
    ExcelObj variantToExcelObj(
      const VARIANT& variant, 
      bool allowRange = false, 
      bool trimArray = true);

    bool trimmedVariantArrayBounds(const VARIANT& variant, size_t& nRows, size_t& nCols);

    VARIANT stringToVariant(const char* str);
    VARIANT stringToVariant(const wchar_t* str);
    VARIANT stringToVariant(const std::wstring_view& str);
  }
}
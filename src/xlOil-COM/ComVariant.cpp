#include "ComVariant.h"
#include <xlOil/TypeConverters.h>
#include <xlOil/ExcelArray.h>
#include <xlOil/Range.h>
#include <xloil/ExcelRef.h>
#include <xloil/ArrayBuilder.h>
#include <xlOil/ExcelTypeLib.h>
#include <xlOil/AppObjects.h>

using std::shared_ptr;
using std::unique_ptr;

namespace xloil
{
  namespace COM
  {
    namespace
    {
      // Small helper function for array conversion
      template<class T> 
      auto elementConvert(const T& val)       { return val; }
      template<>
      auto elementConvert(const VARIANT& val) { return variantToExcelObj(val, false); }

      template<typename T>
      void addStringLength(size_t&, const T&) {}

      template<>
      void addStringLength(size_t& len, const BSTR& v)
      {
        len += wcslen(v);
      }

      template<>
      void addStringLength(size_t& len, const VARIANT& v)
      {
        if (v.vt == VT_BSTR)
          len += wcslen(v.bstrVal);
      }

      auto variantErrorToCellError(SCODE scode)
      {
        return (CellError)(scode - 0x800A07D0);
      }

      bool isNonEmpty(VARIANT& obj)
      {
        switch (obj.vt)
        {
        case VT_BSTR: return wcslen(obj.bstrVal) > 0;
        case VT_ERROR:
          return obj.scode != DISP_E_PARAMNOTFOUND
            && variantErrorToCellError(obj.scode) != CellError::NA;
        case VT_EMPTY: return false;
        }
        return true;
      }

      void trimmedArraySize(VARIANT* data, size_t& nRows, size_t& nCols)
      {
        const auto start = data;
        const auto rows = nRows;

        auto p = start + nCols * nRows - 1; // Go to last element
        for (; nCols > 0; --nCols)
          for (int r = (int)nRows - 1; r >= 0; --r, --p)
            if (isNonEmpty(*p))
              goto StartRowSearch;
       
      StartRowSearch:
       
        for (; nRows > 0; --nRows)
          for (p = start + nRows - 1; p < (start + nCols * nRows); p += rows)
            if (isNonEmpty(*p))
              goto SearchDone;

      SearchDone:;
      }

      template<typename T>
      size_t stringLength(const SafeArrayAccessor<T>& array)
      {
        size_t len = 0;
        for (auto j = 0u; j < array.cols; j++)
          for (auto i = 0u; i < array.rows; i++)
            addStringLength(len, array(i, j));
        return len;
      }
    }

    detail::SafeArrayAccessorBase::SafeArrayAccessorBase(SAFEARRAY* pArr)
      : _ptr(pArr)
      , dimensions(pArr->cDims)
      , cols(dimensions == 1 ? 1 : pArr->rgsabound[0].cElements)
      , rows(pArr->rgsabound[dimensions == 1 ? 0 : 1].cElements)
    {
      if (S_OK != SafeArrayAccessData(pArr, &_data))
        XLO_THROW("Failed to access SafeArray");
    }

    detail::SafeArrayAccessorBase::~SafeArrayAccessorBase()
    {
      SafeArrayUnaccessData(_ptr);
    }

    std::pair<size_t, size_t> SafeArrayAccessor<VARIANT>::trimmedSize() const
    {
      auto nrows = this->rows;
      auto ncols = this->cols;
      trimmedArraySize(data(), nrows, ncols);
      return std::pair(nrows, ncols);
    }

    template<class T>
    auto toExcelObj(SafeArrayAccessor<T> array, bool trimArray)
    {
      if (array.dimensions > 2)
        XLO_THROW("Can only convert 1 or 2 dim arrays");

      auto rows = array.rows;
      auto cols = array.cols;
      if (trimArray)
        std::tie(rows, cols) = array.trimmedSize();

      // We need up-front total string length for the ExcelArrayBuilder
      const auto strLength = stringLength<T>(array);
      ExcelArrayBuilder builder(
        (ExcelObj::row_t)rows, (ExcelObj::col_t)cols, strLength);

      for (auto i = 0u; i < rows; i++)
        for (auto j = 0u; j < cols; j++)
          builder(i, j) = elementConvert(array(i, j));

      return builder.toExcelObj();
    }

    class ToVariant : public ExcelValVisitor<VARIANT>
    {
    public:
      using result_t = return_type;

      using ExcelValVisitor::operator();

      result_t operator()(int x) const
      {
        return _variant_t(x).Detach();
      }
      result_t operator()(bool x) const
      {
        return _variant_t(x).Detach();
      }
      result_t operator()(double x) const
      {
        return _variant_t(x).Detach();
      }
      result_t operator()(const ArrayVal& obj) const
      {
        // No array trimming, for some good reason
        return operator()(ExcelArray(obj, false));
      }
      result_t operator()(const ExcelArray& arr) const
      {
        const auto nRows = arr.nRows();
        const auto nCols = arr.nCols();

        SAFEARRAYBOUND bounds[2];
        bounds[0].lLbound = 0;
        bounds[0].cElements = nRows;
        bounds[1].lLbound = 0;
        bounds[1].cElements = nCols;

        auto array = unique_ptr<SAFEARRAY, HRESULT(__stdcall *)(SAFEARRAY*)> (
          SafeArrayCreate(VT_VARIANT, 2, bounds), SafeArrayDestroy);

        SafeArrayAccessor<VARIANT> arrayData(array.get());

        for (auto i = 0u; i < nRows; i++)
          for (auto j = 0u; j < nCols; j++)
            arrayData(i, j) = (*this)(arr(i, j));

        VARIANT result;
        result.vt = VT_VARIANT | VT_ARRAY;
        result.parray = array.release();

        return result;
      }
      result_t operator()(const PStringRef& pstr) const
      {
        VARIANT result;
        VariantInit(&result);
        V_VT(&result) = VT_BSTR;
        V_BSTR(&result) = SysAllocStringLen(pstr.pstr(), (UINT)pstr.length());
        return result;
      }
      result_t operator()(CellError x) const
      {
        // Magical constant from: 
        // https://docs.microsoft.com/en-us/office/client-developer/excel/how-to-access-dlls-in-excel
        return _variant_t((long)x + (long)0x800A07D0, VT_ERROR).Detach();
      }
      result_t operator()(const RefVal& ref) const
      {
        return operator()(ExcelRef(ref).value());
      }

      // Not part of the usual FromExcel interface, just to aid cascading
      result_t operator()(const ExcelObj& obj) const
      {
        return obj.visit(ToVariant());
      }
    };

    class ToVariantWithRange : public ToVariant
    {
    public:
      using ToVariant::operator();

      result_t operator()(const RefVal& ref) const
      {
        const auto range = ExcelRange(ExcelRef(ref));
        return _variant_t(&range.com()).Detach();
      }
    };

    void excelObjToVariant(VARIANT* v, const ExcelObj& obj, bool allowRange)
    {
      VariantClear(v);
      *v = allowRange
        ? obj.visit(ToVariantWithRange(), vtMissing)
        : obj.visit(ToVariant(), vtMissing);
    }

    ExcelObj variantToExcelObj(const VARIANT& variant, bool allowRange, bool trimArray)
    {
      switch (variant.vt)
      {
      case VT_I2:   return ExcelObj(variant.iVal);
      case VT_I4:   return ExcelObj(variant.lVal);
      case VT_I8:   return ExcelObj(variant.llVal);
      case VT_INT:  return ExcelObj(variant.intVal);
      case VT_UI2:  return ExcelObj(variant.uiVal);
      case VT_UI4:  return ExcelObj(variant.ulVal);
      case VT_UI8:  return ExcelObj(variant.ullVal);
      case VT_UINT: return ExcelObj(variant.uintVal);
      case VT_R8:   return ExcelObj(variant.dblVal);
      case VT_BOOL: return ExcelObj(variant.boolVal == VARIANT_TRUE);
      case VT_BSTR: return ExcelObj((const wchar_t*)variant.bstrVal);
      case VT_CY:   return ExcelObj(variant.cyVal.int64); // currency
      case VT_DATE: return ExcelObj(variant.date);
      case VT_DISPATCH:
      {
        Excel::RangePtr pRange(variant.pdispVal);
        
        // TODO: converting to a ref via the address is a bit expensive
        // if end users (e.g. python) can use an ExcelRange
        auto xlRef = ExcelRef(pRange->GetAddress(
          VARIANT_TRUE, VARIANT_TRUE, Excel::xlA1, VARIANT_TRUE));

        if (allowRange)
          return ExcelObj(std::move(xlRef));
        else
          // Probably faster than variantToExcelObj(pRange->Value2).
          return xlRef.value(); 
      }
      case VT_ERROR:
        return variant.scode == DISP_E_PARAMNOTFOUND
          ? ExcelObj(ExcelType::Missing)
          : ExcelObj(variantErrorToCellError(variant.scode));
      case VT_EMPTY: 
        return ExcelObj();
      }

      if ((variant.vt & VT_ARRAY) == 0)
        XLO_THROW("Unknown variant type {0}", variant.vt);
      else
      {
        auto pArr = variant.parray;
        VARTYPE vartype;
        SafeArrayGetVartype(pArr, &vartype);
        switch (vartype)
        {
        case VT_R8:    return toExcelObj(SafeArrayAccessor<double>(pArr), trimArray);
        case VT_BOOL:  return toExcelObj(SafeArrayAccessor<bool>(pArr), trimArray);
        case VT_BSTR:  return toExcelObj(SafeArrayAccessor<BSTR>(pArr), trimArray);
        case VT_ERROR: return toExcelObj(SafeArrayAccessor<long>(pArr), trimArray);
        case VT_VARIANT: return toExcelObj(SafeArrayAccessor<VARIANT>(pArr), trimArray);
        default:
          XLO_THROW("Unhandled array data type: {0}", variant.vt ^ VT_ARRAY);
        }
      }
    }
    bool trimmedVariantArrayBounds(const VARIANT& variant, size_t& nRows, size_t& nCols)
    {
      if ((variant.vt & VT_ARRAY) == 0)
        return false;

      auto pArr = variant.parray;
      VARTYPE vartype;
      SafeArrayGetVartype(pArr, &vartype);
      if (vartype != VT_VARIANT)
        return false;

      std::tie(nRows, nCols) = SafeArrayAccessor<VARIANT>(pArr).trimmedSize();
      return true;
    }

    VARIANT stringToVariant(const char* str)
    {
      return _variant_t(str).Detach();
    }

    VARIANT stringToVariant(const wchar_t* str)
    {
      return _variant_t(str).Detach();
    }

    VARIANT stringToVariant(const std::wstring_view& str)
    {
      VARIANT result;
      V_VT(&result) = VT_BSTR;
      V_BSTR(&result) = SysAllocStringLen(str.data(), (UINT)str.length());
      return result;
    }
  }
}
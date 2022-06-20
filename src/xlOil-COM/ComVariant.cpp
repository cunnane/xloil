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

        VARIANT* data = nullptr;
        if (S_OK != SafeArrayAccessData(array.get(), (void**)&data))
          XLO_THROW("Failed to access SafeArray");
        
        for (auto i = 0u; i < nRows; i++)
        {
          for (auto j = 0u; j < nCols; j++)
          {
            const auto idx = j * nRows + i;
            auto element = (*this)(arr(i, j));
            data[idx] = element;
          }
        }

        SafeArrayUnaccessData(array.get());

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

    // Small helper function for array conversion
    template<class T> auto elementConvert(const T& val) { return val; }
    ExcelObj elementConvert(const VARIANT& val) { return variantToExcelObj(val, false); }

    template<class T>
    size_t stringLength(T* /*arr*/, size_t /*nRows*/, size_t /*nCols*/)
    {
      return 0;
    }
    template<>
    size_t stringLength<BSTR>(BSTR* pData, size_t nRows, size_t nCols)
    {
      size_t len = 0u;
      for (auto i = 0u; i < nRows; i++)
        for (auto j = 0u; j < nCols; j++)
          len += wcslen(pData[i * nCols + j]);
      return len;
    }
    template<>
    size_t stringLength<VARIANT>(VARIANT* pData, size_t nRows, size_t nCols)
    {
      size_t len = 0;
      for (auto i = 0u; i < nRows; i++)
        for (auto j = 0u; j < nCols; j++)
        {
          auto& p = pData[i * nCols + j];
          if (p.vt == VT_BSTR)
            len += wcslen(p.bstrVal);
        }
      return len;
    }

    template<class T>
    auto arrayToExcelObj(void* pVoidData, size_t nRows, size_t nCols)
    {
      auto pData = (T*)pVoidData;
      const auto strLength = stringLength<T>(pData, nRows, nCols);
      ExcelArrayBuilder builder(
        (ExcelObj::row_t)nRows, (ExcelObj::col_t)nCols, strLength);

      for (auto i = 0u; i < nRows; i++)
        for (auto j = 0u; j < nCols; j++)
        {
          builder(i, j) = elementConvert(pData[j * nRows + i]);
        }
      
      return builder.toExcelObj();
    }

    ExcelObj variantToExcelObj(const VARIANT& variant, bool allowRange)
    {
      switch (variant.vt)
      {
      case VT_R8:   return ExcelObj(variant.dblVal);
      case VT_BOOL: return ExcelObj(variant.boolVal == VARIANT_TRUE);
      case VT_BSTR: return ExcelObj((const wchar_t*)variant.bstrVal);
      case VT_CY:   return ExcelObj(variant.cyVal.int64); // currency
      case VT_DATE: return ExcelObj(variant.date);
      case VT_DISPATCH:
      {
        Excel::Range* pRange;
        if (S_OK != variant.pdispVal->QueryInterface(&pRange))
          XLO_THROW("Unexpected variant type: could not convert to Range");
        
        auto xlRef = ExcelRef(pRange->GetAddress(VARIANT_TRUE, VARIANT_TRUE, Excel::xlA1, VARIANT_TRUE));
        variant.pdispVal->Release(); //TODO: surely pRange->Release();
        
        if (allowRange)
          return xlRef;
        else
          // Probably faster than variantToExcelObj(pRange->Value2).
          return xlRef.value(); 
      }
      case VT_ERROR:
        return variant.scode == DISP_E_PARAMNOTFOUND
          ? ExcelObj(ExcelType::Missing)
          : ExcelObj((CellError)(variant.scode - 0x800A07D0));
      case VT_EMPTY: return ExcelObj();
      }

      if ((variant.vt & VT_ARRAY) == 0)
        XLO_THROW("Unknown variant type {0}", variant.vt);
      else
      {
        const auto dims = variant.parray->cDims;
        if (dims > 2)
          XLO_THROW("Can only convert 1 or 2 dim arrays");

        void* pData;
        if (FAILED(SafeArrayAccessData(variant.parray, &pData)))
          XLO_THROW("Failed calling SafeArrayAccessData");

        std::shared_ptr<SAFEARRAY> pArr(variant.parray, SafeArrayUnaccessData);

        // The rgsabound structure is reversed
        const auto nCols = dims == 1 ? 1 : pArr->rgsabound[0].cElements;
        const auto nRows = pArr->rgsabound[dims == 1 ? 0 : 1].cElements;

        VARTYPE vartype;
        SafeArrayGetVartype(pArr.get(), &vartype);
        switch (vartype)
        {
        case VT_R8:    return arrayToExcelObj<double>(pData, nRows, nCols);
        case VT_BOOL:  return arrayToExcelObj<bool>(pData, nRows, nCols);
        case VT_BSTR:  return arrayToExcelObj<BSTR>(pData, nRows, nCols);
        case VT_ERROR: return arrayToExcelObj<long>(pData, nRows, nCols);
        case VT_VARIANT: return arrayToExcelObj<VARIANT>(pData, nRows, nCols);
        default:
          XLO_THROW("Unhandled array data type: {0}", variant.vt ^ VT_ARRAY);
        }
      }
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
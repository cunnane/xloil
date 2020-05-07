#include "ComVariant.h"
#include <xlOil/TypeConverters.h>
#include <xlOil/ExcelArray.h>
#include <xlOil/ExcelRange.h>
#include "ExcelTypeLib.h"

namespace xloil
{
  namespace COM
  {
    class ToVariant : public FromExcelBase<_variant_t, ToVariant>
    {
    public:
      using result_t = VARIANT;
      result_t fromInt(int x) const
      {
        return _variant_t(x);
      }
      result_t fromBool(bool x) const
      {
        return _variant_t(x);
      }
      result_t fromDouble(double x) const
      {
        return _variant_t(x);
      }
      result_t fromArray(const ExcelObj& obj) const
      {
        return fromArrayObj(ExcelArray(obj, false));
      }
      result_t fromArrayObj(const ExcelArray& arr) const
      {
        // TODO: any benefit from using concrete types for homogenous array?
        // Given this is only used in event handling macros...?
        VARIANT result;
        result.vt = VT_ARRAY | VT_VARIANT;

        const auto nRows = arr.nRows();
        const auto nCols = arr.nCols();

        SAFEARRAYBOUND bounds[2];
        bounds[0].lLbound = 0; bounds[0].cElements = nRows;
        bounds[1].lLbound = 0; bounds[1].cElements = nCols;
        result.parray = SafeArrayCreate(VT_VARIANT, 2, bounds);
        if (!result.parray)
          error();

        VARIANT element;
        for (auto i = 0u; i < nRows; ++i)
        {
          for (auto j = 0u; j < nCols; ++j)
          {
            element = (*this)(arr(i, j));
            long index[] = { i, j };
            SafeArrayPutElement(result.parray, index, &element);
          }
        }
        return result;
      }
      result_t fromString(const wchar_t* buf, size_t len) const
      {
        _variant_t result;
        V_VT(&result) = VT_BSTR;
        V_BSTR(&result) = SysAllocStringLen(buf, (UINT)len);
        return result;
      }
      result_t fromError(CellError x) const
      {
        // Magical constant from: 
        // Excel Add-in Development in C / C++, 2nd Edition by Steve Dalton
        return _variant_t((long)x + (long)2148141008, VT_ERROR);
      }
      result_t fromEmpty(const result_t*) const
      {
        return _variant_t();
      }
      result_t fromMissing(const result_t*) const
      {
        return _variant_t();
      }
      result_t fromRef(const ExcelObj& obj) const
      {
        return fromRef(ExcelRange(obj));
      }
      result_t fromRef(const ExcelRange& r) const
      {
        return (*this)(r.value());
      }
    };

    VARIANT excelObjToVariant(const ExcelObj& obj)
    {
      return ToVariant()(obj);
    }
  }
}
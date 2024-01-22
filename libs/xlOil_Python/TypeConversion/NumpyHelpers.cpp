#include "CPython.h"

#define PY_ARRAY_UNIQUE_SYMBOL xloil_PyArray_API
#define NPY_NO_DEPRECATED_API NPY_1_7_API_VERSION
#include <numpy/arrayobject.h>

#include "NumpyHelpers.h"
#include "NumpyDatetime.h"
#include "Numpy.h"
#include <xloil/Date.h>

typedef npy_int64 npy_datetime;


namespace xloil
{
  namespace Python
  {
    bool importNumpy()
    {
      auto ret = _import_array();
      return ret == 0;
    }

    namespace
    {
      double excelDateFromNumpyDate(const npy_datetime x, const PyArray_DatetimeMetaData& meta)
      {
        npy_datetimestruct dt;
        NpyDatetime_ConvertDatetime64ToDatetimeStruct(
          const_cast<PyArray_DatetimeMetaData*>(&meta), x, &dt);

        if (dt.year == NPY_DATETIME_NAT)
          return NPY_NAN;

        if (meta.base <= NPY_FR_D)
          return excelSerialDateFromYMD((int)dt.year, dt.month, dt.day);
        else
          return excelSerialDateFromYMDHMS(
            (int)dt.year, dt.month, dt.day, dt.hour, dt.min, dt.sec, dt.us);
      }
    }
    PyArray_Descr*
      createDatetimeDtype(int type_num, NPY_DATETIMEUNIT unit)
    {
      PyArray_DatetimeMetaData meta{ unit, 1 };
      return create_datetime_dtype(type_num, &meta);
    }

    template<>
    npy_datetime convertDateTime<NPY_FR_us>(const npy_datetimestruct& dt) noexcept
    {
      PyArray_DatetimeMetaData meta{ NPY_FR_us, 1 };
      npy_datetime result;
      NpyDatetime_ConvertDatetimeStructToDatetime64(&meta, &dt, &result);
      return result;
    }

    FromArrayImpl<NPY_DATETIME>::FromArrayImpl(PyArrayObject* pArr)
      : _meta(get_datetime_metadata_from_dtype(PyArray_DESCR(pArr)))
    {}

    ExcelObj FromArrayImpl<NPY_DATETIME>::toExcelObj(
      ExcelArrayBuilder& /*builder*/,
      void* arrayPtr) const
    {
      auto x = (npy_datetime*)arrayPtr;
      const auto serial = excelDateFromNumpyDate(*x, *_meta);
      return ExcelObj(serial);
    }
  }
}
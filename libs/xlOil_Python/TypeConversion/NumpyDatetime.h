#pragma once

#include "NumpyHelpers.h"

// Prior to the as-yet-unreleased Numpy 2, there are no working API functions
// to convert to and from numpy's datetime representation. There are
// promising looking API functions, but they give a ten-year-old deprecation
// error. The only approach is to copy/paste the relevant conversion code
// which is what we have done here

int
NpyDatetime_ConvertDatetime64ToDatetimeStruct(
  PyArray_DatetimeMetaData* meta, npy_datetime dt,
  npy_datetimestruct* out);

PyArray_Descr*
create_datetime_dtype(int type_num, PyArray_DatetimeMetaData* meta);

int
NpyDatetime_ConvertDatetimeStructToDatetime64(PyArray_DatetimeMetaData* meta,
  const npy_datetimestruct* dts,
  npy_datetime* out);

PyArray_DatetimeMetaData*
get_datetime_metadata_from_dtype(PyArray_Descr* dtype);

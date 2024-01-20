#pragma once

/*
 * Functions to hide the horror of the numpy API, in particular the need to initialise
 * an array of function points in every cpp file. 
 */

#include "CPython.h"
#include <xlOil/ExcelObj.h>
#include <memory>

namespace xloil { 
  class FPArray;
  namespace Python {
    class IPyFromExcel;
} }

namespace xloil
{
  namespace Python
  {
    bool importNumpy();
    bool isArrayDataType(PyTypeObject* p);
    bool isNumpyArray(PyObject* p);

    IPyFromExcel* createFPArrayConverter();

    std::shared_ptr<FPArray> numpyToFPArray(const PyObject& obj);

    PyObject* excelArrayToNumpyArray(const ExcelArray& arr, int dims = 2, int dtype = -1);

    ExcelObj numpyArrayToExcel(const PyObject* p);

    PyObject* toNumpyDatetimeFromExcelDateArray(const PyObject* p);
  }
}
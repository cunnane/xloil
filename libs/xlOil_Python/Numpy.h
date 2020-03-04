#pragma once

/*
 * Functions to hide the horror of the numpy API, in particular the need to initialise
 * an array of function points in every cpp file. 
 */

#include "PyHelpers.h"
#include "ExcelObj.h"

namespace xloil
{
  namespace Python
  {
    bool importNumpy();
    bool isArrayDataType(PyTypeObject* p);
    bool isNumpyArray(PyObject* p);

    PyObject* excelArrayToNumpyArray(const ExcelArray& arr, int dims = 2, int dtype = -1);
    ExcelObj numpyArrayToExcel(const PyObject* p);
  }
}
#pragma once
#include <Python.h>
#include <xlOil/ExcelObj.h>

namespace xloil
{
  namespace Python
  {
    void importDatetime();

    bool isPyDate(PyObject* p);

    ExcelObj pyDateToExcel(PyObject* p);
  }
}
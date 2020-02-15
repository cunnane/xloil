#include <Python.h>
#include "ExcelObj.h"

namespace xloil
{
  namespace Python
  {
    void importDatetime();

    bool isPyDate(PyObject* p);

    ExcelObj pyDateToExcel(PyObject* p);
  }
}
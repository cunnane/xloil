#pragma once

#include <xlOil/ExcelObj.h>
#include "CPython.h"

namespace xloil
{
  namespace Python
  {
    class IPyFromExcel : public IConvertFromExcel<PyObject*>
    {
    public:
      /// <summary>
      /// A useful name for the converter, typically the type supported.
      /// Currently used only for log diagnostics.
      /// </summary>
      virtual const char* name() const = 0;
    };
    class IPyToExcel : public IConvertToExcel<PyObject>
    {
    public:
      /// <summary>
      /// A useful name for the converter, typically the type supported.
      /// Currently used only for log diagnostics.
      /// </summary>
      virtual const char* name() const = 0;
    };
  }
}

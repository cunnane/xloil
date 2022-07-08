#pragma once
#include "PyHelpers.h"

struct IDispatch;

namespace Excel 
{ 
  struct _Application; struct Window; struct _Workbook; struct _Worksheet; struct Range; 
}

namespace xloil
{
  namespace Python
  {
    pybind11::object comToPy(Excel::_Application& p, const char* comlib =nullptr);
    pybind11::object comToPy(Excel::Window& p, const char* comlib = nullptr);
    pybind11::object comToPy(Excel::_Workbook& p, const char* comlib = nullptr);
    pybind11::object comToPy(Excel::_Worksheet& p, const char* comlib = nullptr);
    pybind11::object comToPy(Excel::Range& p, const char* comlib = nullptr);
    pybind11::object comToPy(IDispatch& p, const char* comlib = nullptr);
  }
}

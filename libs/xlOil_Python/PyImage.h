#pragma once
#include <pybind11/pybind11.h>

struct IPictureDisp;

namespace xloil
{
  namespace Python
  {
    IPictureDisp* pictureFromPilImage(const pybind11::object& image);
  }
}
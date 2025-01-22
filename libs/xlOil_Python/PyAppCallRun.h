#pragma once

#include <pybind11/pybind11.h>
#include <string>

namespace xloil { class Application; }

namespace xloil
{
  namespace Python
  {
    pybind11::object applicationRun(
      Application& app, const std::wstring& func, const pybind11::args& args);
  }
}

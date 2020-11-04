#pragma once

namespace xloil
{
  class ExcelObj;

  namespace Python
  {
    class PyFuncInfo;

    void pythonAsyncCallback(
      const PyFuncInfo* info,
      const ExcelObj** xlArgs) noexcept;

    ExcelObj* pythonRtdCallback(
      const PyFuncInfo* info,
      const ExcelObj** xlArgs) noexcept;
  }
}
#pragma once

namespace xloil
{
  class ExcelObj;

  namespace Python
  {
    class PyFuncInfo;

    void pythonAsyncCallback(
      PyFuncInfo* info,
      const ExcelObj* asyncHandle,
      const ExcelObj** xlArgs) noexcept;

    ExcelObj* pythonRtdCallback(
      PyFuncInfo* info,
      const ExcelObj** xlArgs) noexcept;
  }
}
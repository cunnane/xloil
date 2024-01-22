#pragma once
#include <functional>

namespace Excel { struct _Application; }
namespace xloil { class ExcelObj; }

namespace xloil
{
  /// <summary>
  /// Calling XLL specific functions is generally not allowed unless you
  /// are on the main thread and are in an XLL function called by Excel.
  /// This function runs the supplied function object under that context.
  /// Should only be called from the main thread.
  /// </summary>
  bool runInXllContext(const std::function<bool()>& f);

  /// <summary>
  /// Calling XLL specific functions is generally not allowed unless you
  /// are on the main thread and are in an XLL function called by Excel.
  /// This function runs Excel12v on the supplied arguments in that context.
  /// Should only be called from the main thread.
  /// </summary>
  int runInXllContext(
    int func, ExcelObj* result, int nArgs, const ExcelObj** args);
}
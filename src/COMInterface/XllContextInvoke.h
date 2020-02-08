#pragma once
#include <functional>

namespace Excel { struct _Application; }
namespace xloil { class ExcelObj; }

namespace xloil
{
  /// <summary>
  /// Having this class in scope declares that you are on the main thread 
  /// and are in an XLL function called by Excel.
  /// </summary>
  class ScopeInXllContext
  {
  public:
    ScopeInXllContext();
    ~ScopeInXllContext();
    static bool check();
  private:
    static int _count;
  };

  /// <summary>
  /// Calling XLL specific functions is generally not allowed unless you
  /// are on the main thread and are in an XLL function called by Excel.
  /// This function runs the supplied function object under that context.
  /// </summary>
  bool runInXllContext(const std::function<void()>& f);

  /// <summary>
  /// Calling XLL specific functions is generally not allowed unless you
  /// are on the main thread and are in an XLL function called by Excel.
  /// This function runs Excel12v on the supplied arguments in that context.
  /// </summary>
  int runInXllContext(int func, ExcelObj* result, int nArgs, const ExcelObj** args);
}
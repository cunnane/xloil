#include <functional>

namespace Excel { struct _Application; }
namespace xloil { class ExcelObj; }

namespace xloil
{
  Excel::_Application& excelApp();

  class ScopeInXllContext
  {
  public:
    ScopeInXllContext();
    ~ScopeInXllContext();
    static bool check();
  private:
    static int _count;
  };

  bool runInXllContext(const std::function<void()>& f);

  int runInXllContext(int func, ExcelObj* result, int nArgs, const ExcelObj** args);
}
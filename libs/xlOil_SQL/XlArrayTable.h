#include <memory>

struct sqlite3_module;
namespace xloil { class ExcelArray; class ExcelRange; class ExcelObj;  }

namespace xloil
{
  namespace SQL
  {
    using XlArrayInput = ExcelArray;
    using XlRangeInput = ExcelRange;

    extern sqlite3_module XlArrayModule;
    extern sqlite3_module XlRangeModule;
  }
}
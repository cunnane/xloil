#include <memory>

struct sqlite3_module;
namespace xloil { class ExcelArray; class ExcelRef; class ExcelObj;  }

namespace xloil
{
  namespace SQL
  {
    using XlArrayInput = ExcelArray;
    using XlRangeInput = ExcelRef;

    extern sqlite3_module XlArrayModule;
    extern sqlite3_module XlRangeModule;
  }
}
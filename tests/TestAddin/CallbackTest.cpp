#include <xloil/ExcelTypeLib.h>
#include <xloil/StaticRegister.h>
#include <xloil/ExcelCall.h>
#include <xloil/ExcelObj.h>
#include <xloil/ExcelRef.h>
#include <xloil/Date.h>
#include <xloil/ExcelApp.h>

using std::shared_ptr;

namespace xloil
{
  // Replicates the TODAY() function by changing the format of the calling
  // cell to date
  XLO_FUNC_START(testToday())
  {
    CallerInfo caller;
    if (!caller.fullSheetName().empty())
      excelPost([=] 
      {
        excelApp().Range[caller.writeAddress().c_str()]->NumberFormat = L"dd-mm-yyyy"; 
      });
    std::tm buf; 
    auto now = std::time(0);
    localtime_s(&buf, &now);
    return returnValue(buf);
  }
  XLO_FUNC_END(testToday);
}
#include <xloil/ExcelTypeLib.h>
#include <xloil/StaticRegister.h>
#include <xloil/ExcelCall.h>
#include <xloil/ExcelObj.h>
#include <xloil/Events.h>
#include <xloil/ExcelRef.h>
#include <xloil/Date.h>
#include <xloil/ExcelThread.h>
#include <xloil/AppObjects.h>

using std::shared_ptr;

namespace xloil
{
  // Replicates the TODAY() function by changing the format of the calling
  // cell to date
  XLO_FUNC_START(testToday())
  {
    CallerInfo caller;
    if (!caller.fullSheetName().empty())
    {
      auto handle = xloil::Event::SheetChange().bind(
        [=](const wchar_t* wsName, const Range& /*target*/)
        {
          // Could check range here as well to avoid
          if (wsName == caller.sheetName())
            thisApp().com().Range[caller.address().c_str()]->NumberFormat = L"dd-mm-yyyy";
        }
      );
      auto milliSecsDelay = 1000;
      runExcelThread([=]() mutable
      {
        handle.reset(); // Removes the SheetChange handler
      }, ExcelRunQueue::WINDOW, milliSecsDelay);
    }
    std::tm buf; 
    auto now = std::time(0);
    localtime_s(&buf, &now);
    return returnValue(buf);
  }
  XLO_FUNC_END(testToday);
}
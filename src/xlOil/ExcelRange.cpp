#include <xloil/ExcelRange.h>
#include <xlOil/ExcelRef.h>
#include <Cominterface/ComRange.h>
#include <Cominterface/XllContextInvoke.h>

namespace xloil
{
  Range* newXllRange(const ExcelObj& xlRef)
  {
    return new XllRange(xlRef);
  }
  Range* newRange(const wchar_t* address)
  {
    if (InXllContext::check())
      return new XllRange(ExcelRef(address));
    else
      return new COM::ComRange(address);
  }
}
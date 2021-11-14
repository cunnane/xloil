#include <xloil/ExcelRange.h>
#include <xlOil/ExcelRef.h>
#include <xlOil/AppObjects.h>
#include <xlOil-COM/XllContextInvoke.h>

namespace xloil
{
  Range* newRange(const wchar_t* address)
  {
    if (InXllContext::check())
      return new XllRange(ExcelRef(address));
    else
      return new ExcelRange(address);
  }
}
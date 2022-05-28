#include <xloil/AppObjects.h>
#include <xlOil/ExcelThread.h>
#include <xlOil/ExcelTypeLib.h>
#include <xlOil/WindowsSlim.h>

#include <xlOil-COM/Connect.h>
#include <xlOil-COM/ComAddin.h>
#include <xlOil-COM/ComVariant.h>
#include <xloil/AppObjects.h>
#include <xloil/Log.h>
#include <xloil/Throw.h>
#include <xloil/State.h>
#include <xloil/ExcelUI.h>
#include <comdef.h>
using std::make_shared;
using std::shared_ptr;
using std::vector;
using std::wstring;

namespace xloil
{
  std::shared_ptr<IComAddin> xloil::makeComAddin(
    const wchar_t* name, const wchar_t* description)
  {
    return COM::createComAddin(name, description);
  }

  ExcelObj variantToExcelObj(const VARIANT& variant, bool allowRange)
  {
    return COM::variantToExcelObj(variant, allowRange);
  }

  void excelObjToVariant(VARIANT* v, const ExcelObj& obj)
  {
    COM::excelObjToVariant(v, obj);
  }
  
  void statusBarMsg(const std::wstring_view& msg, size_t timeout)
  {
    if (!msg.empty())
      runExcelThread([msg = wstring(msg)]() { 
        excelApp().com().PutStatusBar(0, msg.c_str()); 
      });
    
    // Send a null str to PutStatusBar in 'timeout' millisecs to clear it
    if (timeout > 0)
      runExcelThread([]() {
          excelApp().com().PutStatusBar(0, _bstr_t()); 
        }, ExcelRunQueue::COM_API, (unsigned)timeout);
  }
}
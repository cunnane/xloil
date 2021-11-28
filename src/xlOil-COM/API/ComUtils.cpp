#include "..\..\..\include\xloil\AppObjects.h"
#include <xlOil/ExcelApp.h>
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

namespace xloil
{
  Excel::_Application& excelApp() noexcept
  {
    return COM::excelApp();
  }

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
  
  ExcelWindow::ExcelWindow(const wchar_t* caption)
  {
    try
    {
      auto winptr = caption
        ? excelApp().Windows->GetItem(caption)
        : excelApp().ActiveWindow;
      init(winptr);
    }
    XLO_RETHROW_COM_ERROR;
  }
  size_t ExcelWindow::hwnd() const
  {
    return (size_t)ptr()->Hwnd;
  }
  std::wstring ExcelWindow::name() const
  {
    return ptr()->Caption.bstrVal;
  }
  ExcelWorkbook ExcelWindow::workbook() const
  {
    try
    {
      return ExcelWorkbook(Excel::_WorkbookPtr(ptr()->Parent));
    }
    XLO_RETHROW_COM_ERROR;
  }
  void statusBarMsg(const std::wstring_view& msg, size_t timeout)
  {
    if (!msg.empty())
      runExcelThread([str = std::wstring(msg)](){excelApp().PutStatusBar(0, str.c_str()); });
    if (timeout > 0)
      runExcelThread([]() { excelApp().PutStatusBar(0, _bstr_t()); }, ExcelRunQueue::COM_API, 10, 200, timeout);
  }

  IAppObject::~IAppObject()
  {
    if (_ptr)
      _ptr->Release();
  }

  void IAppObject::init(IDispatch* ptr)
  {
    _ptr = ptr;
    if (ptr)
      ptr->AddRef();
  }

  void IAppObject::assign(const IAppObject& that)
  {
    if (_ptr) _ptr->Release();
    _ptr = that._ptr;
    _ptr->AddRef();
  }

  ExcelWorkbook::ExcelWorkbook(const wchar_t* name)
  {
    try
    {
      auto ptr = name
        ? excelApp().Workbooks->GetItem(name)
        : excelApp().ActiveWorkbook;
      init(ptr);
    }
    XLO_RETHROW_COM_ERROR;
  }

  std::wstring ExcelWorkbook::name() const
  {
    return ptr()->Name.GetBSTR();
  }

  std::wstring ExcelWorkbook::path() const
  {
    return ptr()->Path.GetBSTR();
  }
  std::vector<ExcelWindow> ExcelWorkbook::windows() const
  {
    try
    {
      vector<ExcelWindow> result;
      for (auto i = 1; i <= ptr()->Windows->Count; ++i)
        result.emplace_back(ptr()->Windows->GetItem(i));
      return result;
    }
    XLO_RETHROW_COM_ERROR;
  }

  void ExcelWorkbook::activate() const
  {
    ptr()->Activate();
  }

  namespace App
  {
    ExcelWorkbook activeWorkbook()
    {
      return ExcelWorkbook();
    }

    std::vector<ExcelWorkbook> workbooks()
    {
      try
      {
        auto& app = excelApp();
        vector<ExcelWorkbook> result;
        for (auto i = 1; i <= app.Workbooks->Count; ++i)
          result.emplace_back(app.Workbooks->GetItem(i));
        return std::move(result);
      }
      XLO_RETHROW_COM_ERROR;
    }

    ExcelWindow activeWindow()
    {
      return ExcelWindow();
    }
    std::vector<ExcelWindow> windows()
    {
      try
      {
        auto& app = excelApp();
        vector<ExcelWindow> result;
        for (auto i = 1; i <= app.Windows->Count; ++i)
          result.emplace_back(app.Windows->GetItem(i));
        return std::move(result);
      }
      XLO_RETHROW_COM_ERROR;
    }
  }
}
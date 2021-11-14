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

  IComAddin* xloil::makeComAddin(
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

  ExcelWindow::ExcelWindow(Excel::Window* ptr) 
    : _window(ptr) 
  {
    _window->AddRef();
  }
  
  ExcelWindow::ExcelWindow(const wchar_t* caption)
  {
    try
    {
      auto winptr = caption
        ? excelApp().Windows->GetItem(caption)
        : excelApp().ActiveWindow;
      _window = winptr;
      _window->AddRef();
    }
    XLO_RETHROW_COM_ERROR;
  }
  size_t ExcelWindow::hwnd() const
  {
    return (size_t)_window->Hwnd;
  }
  std::wstring ExcelWindow::name() const
  {
    return _window->Caption.bstrVal;
  }
  ExcelWorkbook ExcelWindow::workbook() const
  {
    return ExcelWorkbook(Excel::_WorkbookPtr(_window->Parent));
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
    auto p = basePtr();
    if (p)
      p->Release();
  }
  ExcelWorkbook::ExcelWorkbook(Excel::_Workbook* p)
    : _wb(p)
  {
    p->AddRef();
  }

  ExcelWorkbook::ExcelWorkbook(const wchar_t* name)
  {
    try
    {
      auto ptr = name
        ? excelApp().Workbooks->GetItem(name)
        : excelApp().ActiveWorkbook;
      _wb = ptr;
      ptr->AddRef();
    }
    XLO_RETHROW_COM_ERROR;
  }

  std::wstring ExcelWorkbook::name() const
  {
    return _wb->Name;
  }

  std::wstring ExcelWorkbook::path() const
  {
    return _wb->Path;
  }
  std::vector<ExcelWindow> ExcelWorkbook::windows() const
  {
    try
    {
      vector<ExcelWindow> result;
      for (auto i = 0; i < _wb->Windows->Count; ++i)
        result.emplace_back(_wb->Windows->GetItem(i));
      return result;
    }
    XLO_RETHROW_COM_ERROR;
  }

  void ExcelWorkbook::activate() const
  {
    _wb->Activate();
  }

  namespace App
  {
    ExcelWorkbook activeWorkbook()
    {
      return ExcelWorkbook();
    }

    std::vector<ExcelWorkbook> workbooks()
    {
      auto& app = excelApp();
      vector<ExcelWorkbook> result;
      for (auto i = 1; i <= app.Workbooks->Count; ++i)
        result.emplace_back(app.Workbooks->GetItem(i));
      return std::move(result);
    }

    ExcelWindow activeWindow()
    {
      return ExcelWindow();
    }
    std::vector<ExcelWindow> windows()
    {
      auto& app = excelApp();
      vector<ExcelWindow> result;
      for (auto i = 1; i <= app.Windows->Count; ++i)
        result.emplace_back(app.Windows->GetItem(i));
      return std::move(result);
    }

  }
}
#include <xlOil/ExcelApp.h>
#include <xlOil/ExcelTypeLib.h>
#include <xlOil/WindowsSlim.h>
#include <xlOil-COM/Connect.h>
#include <xlOil-COM/ComAddin.h>
#include <xlOil-COM/ComVariant.h>
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
  ExcelWindow::~ExcelWindow()
  {
    if (_window)
      _window->Release();
  }
  size_t ExcelWindow::hwnd() const
  {
    return (size_t)_window->Hwnd;
  }
  std::wstring ExcelWindow::caption() const
  {
    return _window->Caption.bstrVal;
  }
  std::wstring ExcelWindow::workbook() const
  {
    auto wb = Excel::_WorkbookPtr(_window->Parent);
    return wb->Name.GetBSTR();
  }

  std::vector<std::shared_ptr<ExcelWindow>> workbookWindows(const wchar_t* wbName)
  {
    try
    {
      auto wb = wbName ? excelApp().Workbooks->GetItem(wbName) : excelApp().ActiveWorkbook;
      vector<shared_ptr<ExcelWindow>> result(wb->Windows->Count);
      for (auto i = 0; i < result.size(); ++i)
        result[i].reset(new ExcelWindow(wb->Windows->GetItem(i)));
      return result;
    }
    XLO_RETHROW_COM_ERROR;
  }
}
#include <xloil/AppObjects.h>
#include <xlOil/ExcelApp.h>
#include <xlOil/ExcelTypeLib.h>
#include <xlOil/WindowsSlim.h>
#include <xlOil/ExcelRange.h>
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
  namespace
  {
    template <class T>
    struct CollectionToVector
    {
      template <class V>
      vector<T> operator()(const V& collection) const
      {
        try
        {
          vector<T> result;
          const auto N = collection->Count;
          for (auto i = 1; i <= N; ++i)
            result.emplace_back(collection->GetItem(i));
          return std::move(result);
        }
        XLO_RETHROW_COM_ERROR;
      }
    };

    _variant_t stringToVariant(const std::wstring_view& str)
    {
      auto variant = COM::stringToVariant(str);
      return _variant_t(variant, false);
    }
  }
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
  
  ExcelWindow::ExcelWindow(const std::wstring_view& caption)
  {
    try
    {
      if (caption.empty())
        init(excelApp().ActiveWindow);
      else
        init(excelApp().Windows->GetItem(stringToVariant(caption)));
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

  ExcelWorkbook::ExcelWorkbook(const std::wstring_view& name)
  {
    try
    {
      if (name.empty())
        init(excelApp().ActiveWorkbook);
      else
        init(excelApp().Workbooks->GetItem(stringToVariant(name)));
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
    return CollectionToVector<ExcelWindow>()(ptr()->Windows);
  }

  void ExcelWorkbook::activate() const
  {
    ptr()->Activate();
  }

  vector<ExcelWorksheet> ExcelWorkbook::worksheets() const
  {
    try
    {
      vector<ExcelWorksheet> result;
      const auto N = ptr()->Worksheets->Count;
      for (auto i = 1; i <= N; ++i)
        result.push_back((Excel::_Worksheet*)(IDispatch*)ptr()->Worksheets->GetItem(i));
      return std::move(result);
    }
    XLO_RETHROW_COM_ERROR;
  }
  ExcelWorksheet ExcelWorkbook::worksheet(const std::wstring_view& name) const
  {
    try
    {
      return (Excel::_Worksheet*)(IDispatch*)(ptr()->Worksheets->GetItem(stringToVariant(name)));
    }
    XLO_RETHROW_COM_ERROR;
  }

  std::wstring ExcelWorksheet::name() const
  {
    return ptr()->Name.GetBSTR();
  }

  ExcelWorkbook ExcelWorksheet::parent() const
  {
    return ExcelWorkbook((Excel::_Workbook*)(IDispatch*)ptr()->Parent);
  }

  ExcelRange ExcelWorksheet::range(
    int fromRow, int fromCol,
    int toRow, int toCol) const
  {
    try
    {
      if (toRow == Range::TO_END)
        toRow = ptr()->Rows->GetCount();
      if (toCol == Range::TO_END)
        toCol = ptr()->Columns->GetCount();

      auto r = ptr()->GetRange(
        ptr()->Cells->Item[fromRow - 1][fromCol - 1],
        ptr()->Cells->Item[toRow - 1][toCol - 1]);
      return ExcelRange(r);
    }
    XLO_RETHROW_COM_ERROR;
  }

  ExcelRange ExcelWorksheet::range(const std::wstring_view& address) const
  {
    auto fullAddress = std::wstring(ptr()->Name);
    fullAddress += '!';
    fullAddress += address;
    return ExcelRange(fullAddress.c_str());
  }
  ExcelObj ExcelWorksheet::value(Range::row_t i, Range::col_t j) const
  {
    return COM::variantToExcelObj(ptr()->Cells->Item[i][j]);
  }
  void ExcelWorksheet::activate()
  {
    try
    {
      ptr()->Activate();
    }
    XLO_RETHROW_COM_ERROR;
  }
  void ExcelWorksheet::calculate()
  {
    try
    {
      ptr()->Calculate();
    }
    XLO_RETHROW_COM_ERROR;
  }

  namespace App
  {
    ExcelWorkbook Workbooks::active()
    {
      return ExcelWorkbook();
    }
    std::vector<ExcelWorkbook> Workbooks::list()
    {
      return CollectionToVector<ExcelWorkbook>()(excelApp().Workbooks);
    }
    size_t Workbooks::count()
    {
      return excelApp().Workbooks->Count;
    }

    ExcelWindow Windows::active()
    {
      return ExcelWindow();
    }
    std::vector<ExcelWindow> Windows::list()
    {
      return CollectionToVector<ExcelWindow>()(excelApp().Windows);
    }
    size_t Windows::count()
    {
      return excelApp().Windows->Count;
    }

    ExcelWorksheet Worksheets::active()
    {
      try
      {
        Excel::_Worksheet* sheet = nullptr;
        excelApp().ActiveSheet->QueryInterface(&sheet);
        return ExcelWorksheet(sheet);
      }
      XLO_RETHROW_COM_ERROR;
    }
  }
}
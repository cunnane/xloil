#include <xloil/AppObjects.h>
#include <xlOil-COM/Connect.h>
#include <xlOil-COM/ComVariant.h>
#include <xlOil/ExcelTypeLib.h>
#include <xlOil/Range.h>
#include <xloil/Log.h>
#include <xloil/Throw.h>
#include <xloil/State.h>
#include <functional>
#include <comdef.h>

using std::shared_ptr;
using std::make_shared;
using std::vector;
using std::wstring;

namespace xloil
{
  namespace
  {
    template <class T>
    struct CollectionToVector
    {
      using Com_t = decltype(&T(nullptr).com());
      template <class V>
      vector<T> operator()(const V& collection) const
      {
        return operator()(collection, collection->GetCount());
      }
      template <class V>
      vector<T> operator()(const V& collection, const long N) const
      {
        try
        {
          vector<T> result;
          for (auto i = 1; i <= N; ++i)
            result.emplace_back((Com_t)(IDispatch*)collection->GetItem(i));
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

  Application& excelApp()
  {
    return COM::attachedApplication();
  }

  void DispatchObject::release()
  {
    if (_ptr)
    {
      _ptr->Release();
      _ptr = nullptr;
    }
  }

  void DispatchObject::init(IDispatch* ptr, bool steal)
  {
    _ptr = ptr;
    if (!steal && ptr)
      ptr->AddRef();
  }

  Application::Application(Excel::_Application* app)
    : AppObject(app ? app : COM::newApplicationObject(), true)
  {
    if (!valid())
      throw ComConnectException("Failed to create Application object");
  }

  Application::Application(size_t hWnd)
    : AppObject([hWnd]() {
        auto p = COM::applicationObjectFromWindow((HWND)hWnd);
        if (!p)
          throw ComConnectException("Failed to create Application object from window handle");
        return p;
      }(), true)
  {
  }

  namespace
  {
    Excel::_Application* workbookFinder(const wchar_t* workbook)
    {
      HWND xlmain = 0;
      while ((xlmain = COM::nextExcelMainWindow(xlmain)) != 0)
      {
        auto xlApp = Application(COM::applicationObjectFromWindow(xlmain));
        ExcelWorkbook wb(nullptr);
        if (xlApp.workbooks().tryGet(workbook, wb))
          return &xlApp.com();
      }
      return nullptr;
    }
  }

  Application::Application(const wchar_t* workbook)
    : AppObject(workbookFinder(workbook))
  {
    if (!valid())
      throw ComConnectException("Failed to create Application object from workbook");
  }

  std::wstring Application::name() const
  {
    try
    {
      return com().Name.GetBSTR();
    }
    XLO_RETHROW_COM_ERROR;
  }

  void Application::calculate(const bool full, const bool rebuild)
  {
    try
    {
      if (rebuild)
        com().CalculateFullRebuild();
      else if (full)
        com().CalculateFull();
      else
        com().Calculate();
    }
    XLO_RETHROW_COM_ERROR;
  }

  ExcelWorksheet Application::ActiveWorksheet() const
  {
    try
    {
      Excel::_Worksheet* sheet = nullptr;
      com().ActiveSheet->QueryInterface(&sheet);
      return ExcelWorksheet(sheet);
    }
    XLO_RETHROW_COM_ERROR;
  }

  void Application::quit(bool silent)
  {
    try
    {
      if (!valid())
        return;

      if (silent)
        com().PutDisplayAlerts(0, VARIANT_FALSE);
      com().Quit();

      // Release the COM object so app really does quit
      release();
    }
    XLO_RETHROW_COM_ERROR;
  }

  bool Application::getVisible() const 
  {
    try
    {
      return com().Visible;
    }
    XLO_RETHROW_COM_ERROR;
  }

  void Application::setVisible(bool x)
  {
    try
    { 
      com().PutVisible(0, x ? VARIANT_TRUE : VARIANT_FALSE);
    }
    XLO_RETHROW_COM_ERROR;
  }

  bool Application::getEnableEvents()
  {
    try
    {
      return com().EnableEvents;
    }
    XLO_RETHROW_COM_ERROR;
  }

  void Application::setEnableEvents(bool value)
  {
    try
    {
      com().EnableEvents = _variant_t(value);
    }
    XLO_RETHROW_COM_ERROR;
  }

  namespace
  {
    template <typename F, typename T, std::size_t N, std::size_t... Idx>
    decltype(auto) appRun_impl(F func, T(&args)[N], std::index_sequence<Idx...>) {
      return excelApp().com().Run(func, args[Idx]...);
    }

    template <typename T, std::size_t N>
    decltype(auto) appRun(const wchar_t* func, T(&args)[N]) {
      return appRun_impl(func, args, std::make_index_sequence<N>{});
    }
  }

  ExcelObj Application::Run(
    const std::wstring& func,
    const size_t nArgs,
    const ExcelObj* args[])
  {
    if (nArgs > 30)
      XLO_THROW("Application::Run maximum number of args is 30");

    static _variant_t vArgs[30] = {
      vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
      vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
      vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
      vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
      vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
      vtMissing, vtMissing, vtMissing, vtMissing, vtMissing
    };

    // The construction of 'cleanup' is all noexcept
    auto finally = [begin = vArgs, end = vArgs + nArgs](void*)
    {
      for (auto i = begin; i != end; ++i)
        *i = vtMissing;
    };
    std::unique_ptr<void, decltype(finally)> cleanup((void*)1, finally);

    for (size_t i = 0; i < nArgs; ++i)
      COM::excelObjToVariant(&vArgs[i], *args[i], true);

    try
    {
      auto result = appRun(func.c_str(), vArgs);
      return COM::variantToExcelObj(result);
    }
    XLO_RETHROW_COM_ERROR;
  }

  ExcelWorkbook Application::Open(
    const std::wstring& filepath, 
    bool updateLinks, 
    bool readOnly)
  {
    try
    {
      return ExcelWorkbook(com().Workbooks->Open(
        _bstr_t(filepath.c_str()), updateLinks ? 3 : 0, _variant_t(readOnly)
      ).Detach(), true);
    }
    XLO_RETHROW_COM_ERROR;
  }

  ExcelWindow::ExcelWindow(const std::wstring_view& caption, Application app)
    : AppObject([&]() {
        try
        {
          if (caption.empty())
            return app.com().ActiveWindow.Detach();
          else
            return app.com().Windows->GetItem(stringToVariant(caption)).Detach();
        }
        XLO_RETHROW_COM_ERROR;
      }(), true)
  {}

  size_t ExcelWindow::hwnd() const
  {
    return (size_t)com().Hwnd;
  }

  std::wstring ExcelWindow::name() const
  {
    return com().Caption.bstrVal;
  }

  Application ExcelWindow::app() const
  {
    return Application(com().Application.Detach());
  }

  ExcelWorkbook ExcelWindow::workbook() const
  {
    try
    {
      return ExcelWorkbook(Excel::_WorkbookPtr(com().Parent));
    }
    XLO_RETHROW_COM_ERROR;
  }

  ExcelWorkbook::ExcelWorkbook(const std::wstring_view& name, Application app)
    : AppObject([&]() {
        try
        {
          if (name.empty())
            return app.com().ActiveWorkbook.Detach();
          else
          {
            auto workbooks = app.com().Workbooks;
            return workbooks->GetItem(stringToVariant(name)).Detach();
          }
        }
        XLO_RETHROW_COM_ERROR;
      }(), true)
  {}

  std::wstring ExcelWorkbook::name() const
  {
    return com().Name.GetBSTR();
  }

  Application ExcelWorkbook::app() const
  {
    return Application(com().Application.Detach());
  }

  std::wstring ExcelWorkbook::path() const
  {
    return com().Path.GetBSTR();
  }

  void ExcelWorkbook::activate() const
  {
    com().Activate();
  }

  ExcelWorksheet ExcelWorkbook::add(
    const std::wstring_view& name, 
    const ExcelWorksheet& before, 
    const ExcelWorksheet& after) const
  {
    try
    {
      if (before.valid() && after.valid())
        XLO_THROW("ExcelWorkbook::add: at most one of 'before' and 'after' should be specified");

      auto ws = ExcelWorksheet((Excel::_Worksheet*)(com().Worksheets->Add(
        before.valid() ? _variant_t(&before.com()) : vtMissing,
        after.valid() ? _variant_t(&after.com()) : vtMissing).Detach()), true);
      if (!name.empty())
        ws.setName(name);
      return ws;
    }
    XLO_RETHROW_COM_ERROR;
  }

  void ExcelWorkbook::save(const std::wstring_view& filepath)
  {
    try
    {
      if (filepath.empty())
        com().Save();
      else
        com().SaveAs(stringToVariant(filepath), 
          vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, Excel::XlSaveAsAccessMode::xlNoChange);
    }
    XLO_RETHROW_COM_ERROR;
  }

  void ExcelWorkbook::close(bool save)
  {
    try
    {
      com().Close(_variant_t(save));
    }
    XLO_RETHROW_COM_ERROR;
  }

  std::wstring ExcelWorksheet::name() const
  {
    try
    {
      return com().Name.GetBSTR();
    }
    XLO_RETHROW_COM_ERROR;
  }

  Application ExcelWorksheet::app() const
  {
    return Application(com().Application.Detach());
  }

  ExcelWorkbook ExcelWorksheet::parent() const
  {
    try
    {
      return ExcelWorkbook((Excel::_Workbook*)(IDispatch*)com().Parent);
    }
    XLO_RETHROW_COM_ERROR;
  }

  ExcelRange ExcelWorksheet::range(
    int fromRow, int fromCol,
    int toRow, int toCol) const
  {
    try
    {
      if (toRow == Range::TO_END)
        toRow = com().Rows->GetCount();
      if (toCol == Range::TO_END)
        toCol = com().Columns->GetCount();

      auto r = com().GetRange(
        com().Cells->Item[fromRow + 1][fromCol + 1],
        com().Cells->Item[toRow + 1][toCol + 1]);
      return ExcelRange(r);
    }
    XLO_RETHROW_COM_ERROR;
  }

  ExcelRange ExcelWorksheet::range(const std::wstring_view& address) const
  {
    try
    {
      return ExcelRange(formatStr(L"'[%s]%s'!%s", 
        parent().name().c_str(), 
        name().c_str(), 
        wstring(address).data()));
    }
    XLO_RETHROW_COM_ERROR;
  }

  ExcelObj ExcelWorksheet::value(Range::row_t i, Range::col_t j) const
  {
    return COM::variantToExcelObj(com().Cells->Item[i][j]);
  }

  void ExcelWorksheet::activate()
  {
    try
    {
      com().Activate();
    }
    XLO_RETHROW_COM_ERROR;
  }

  void ExcelWorksheet::calculate()
  {
    try
    {
      com().Calculate();
    }
    XLO_RETHROW_COM_ERROR;
  }

  void ExcelWorksheet::setName(const std::wstring_view& name)
  {
    try
    {
      com().Name = stringToVariant(name).bstrVal;
    }
    XLO_RETHROW_COM_ERROR;
  }

  namespace 
  {
    template<class TRes, class TObj>
    auto comGet(const TObj& obj, const std::wstring_view& what)
    {
      try
      {
        using Com_t = std::remove_reference_t<decltype(TRes(nullptr).com())>;
        return TRes((Com_t*)obj->GetItem(stringToVariant(what)).Detach(), true);
      }
      XLO_RETHROW_COM_ERROR;
    }

    template<class TObj, class TRes>
    bool comTryGet(const TObj& obj, const std::wstring_view& what, TRes& out)
    {
      // See other possibility here. Seems a bit crazy?
      // https://stackoverflow.com/questions/9373082/detect-whether-excel-workbook-is-already-open
      try
      {
        using Com_t = std::remove_reference_t<decltype(TRes(nullptr).com())>;
        out = TRes((Com_t*)obj->GetItem(stringToVariant(what)).Detach(), true);
        return true;
      }
      catch (_com_error& error)
      {
        if (error.Error() == DISP_E_BADINDEX)
          return false;
        XLO_THROW(L"COM Error {0:#x}: {1}", (size_t)error.Error(), error.ErrorMessage());
      }
    }
  }

  bool Workbooks::tryGet(const std::wstring_view& workbookName, ExcelWorkbook& wb) const
  {
    return comTryGet(this->app.com().Workbooks, workbookName, wb);
  }

  ExcelWorkbook Workbooks::add()
  {
    try
    {
      auto p = app.com().Workbooks->Add();
      return ExcelWorkbook(p.Detach(), true);
    }
    XLO_RETHROW_COM_ERROR;
  }

  Worksheets::Worksheets(Application app)
    : parent(app.workbooks().active())
  {
    if (!parent.valid())
      XLO_THROW("No active workbook");
  }

  Worksheets::Worksheets(ExcelWorkbook workbook)
    : parent(workbook)
  {}

  vector<ExcelWorksheet> Worksheets::list() const
  {
    try
    {
      // TODO: Not sure why I need to pass the N here... if not, get `unresolved external`
      auto N = parent.com().Worksheets->Count;
      return CollectionToVector<ExcelWorksheet>()(parent.com().Worksheets, N);
    }
    XLO_RETHROW_COM_ERROR;
  }

  ExcelWorksheet Worksheets::get(const std::wstring_view& name) const
  {
    return comGet<ExcelWorksheet>(parent.com().Worksheets, name);
  }
  
  bool Worksheets::tryGet(const std::wstring_view& name, ExcelWorksheet& out) const
  {
    return comTryGet(parent.com().Worksheets, name, out);
  }

  size_t Worksheets::count() const
  {
    return parent.com().Worksheets->GetCount();
  }

  Workbooks::Workbooks(Application app)
    : app(app)
  {}

  ExcelWorkbook Workbooks::active() const
  {
    return ExcelWorkbook(std::wstring_view(), app);
  }

  std::vector<ExcelWorkbook> Workbooks::list() const
  {
    return CollectionToVector<ExcelWorkbook>()(app.com().Workbooks);
  }

  size_t Workbooks::count() const
  {
    return app.com().Workbooks->GetCount();
  }

  Windows::Windows(const Application& app)
    : AppObject(app.com().Windows.Detach(), true)
  {}

  Windows::Windows(const ExcelWorkbook& workbook)
    : AppObject(workbook.com().Windows.Detach(), true)
  {}

  ExcelWindow Windows::active() const
  {
    return ExcelWindow(std::wstring_view(), app());
  }

  Application Windows::app() const
  {
    return Application(com().Application.Detach());
  }

  bool Windows::tryGet(const std::wstring_view& name, ExcelWindow& out) const
  {
    return comTryGet(&com(), name, out);
  }

  std::vector<ExcelWindow> Windows::list() const
  {
    return CollectionToVector<ExcelWindow>()(&com());
  }

  size_t Windows::count() const
  {
    return com().GetCount();
  }
}

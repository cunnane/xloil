#include <xlOil/ExcelTypeLib.h>
#include <xloil/AppObjects.h>
#include <xlOil-COM/Connect.h>
#include <xlOil-COM/ComVariant.h>
#include <xlOil-COM/ComEventSink.h>

#include <xlOil/Range.h>
#include <xloil/Log.h>
#include <xloil/Throw.h>
#include <xloil/State.h>
#include <functional>
#include <comdef.h>
#include <tuple>


using std::shared_ptr;
using std::make_shared;
using std::vector;
using std::wstring;

void workaroundVisualStudioBug()
{
  /*
  As far as COM is concerned, "templates" are a new-fangled trick, to be
  treated with suspicion and comtempt.  Duly, Microsoft's compiler
  declines to look inside templated code for references to function bodies
  which it really ought to have been pulling from the TLI file. Therefore,
  we need to specify which functions we want to use in this nice explicit way
  which should be impossible for even the most simple-minded compiler to get
  wrong.
  */
  ((Excel::Windows*)(nullptr))->Get_NewEnum();
  ((Excel::Windows*)(nullptr))->GetCount();
  ((Excel::Windows*)(nullptr))->GetApplication();
  ((Excel::Windows*)(nullptr))->GetItem(_variant_t());
  ((Excel::Workbooks*)(nullptr))->Get_NewEnum();
  ((Excel::Workbooks*)(nullptr))->GetCount();
  ((Excel::Workbooks*)(nullptr))->GetApplication();
  ((Excel::Sheets*)(nullptr))->Get_NewEnum();
  ((Excel::Sheets*)(nullptr))->GetCount();
  ((Excel::Sheets*)(nullptr))->GetApplication();
  ((Excel::Sheets*)(nullptr))->GetItem(_variant_t());
  ((Excel::Areas*)(nullptr))->Get_NewEnum();
  ((Excel::Areas*)(nullptr))->GetCount();
  ((Excel::Areas*)(nullptr))->GetApplication();
  ((Excel::Areas*)(nullptr))->GetItem(0);
}

namespace xloil
{
  using detail::UnknownObject;
  using detail::AppObject;

  namespace
  {
    template <class T>
    using ComPtr_t = _com_ptr_t<_com_IIID<T, &__uuidof(T)>>;

    template<typename T, class V>
    struct comPtrCast
    {
      auto operator()(const ComPtr_t<V>& v) const { return ComPtr_t<T>(v); }
    };
    template<typename T>
    struct comPtrCast<T, T>
    {
      auto operator()(const ComPtr_t<T>& v) const { return v; }
    };

    template<typename T>
    using ComType = std::remove_reference_t<decltype(T(nullptr).com())>;

    template<typename TAppObj, class V>
    auto fromComPtr(const ComPtr_t<V>& v)
    {
      return TAppObj(comPtrCast<ComType<TAppObj>, V>()(v).Detach(), true);
    }

    template<typename Sig>
    struct signature;

    template<typename R, typename ...Args>
    struct signature<R(Args...)>
    {
      using type = std::tuple<Args...>;
    };

    template<typename C, typename R, typename ...Args>
    struct signature<R (C::*)(Args...)>
    {
      using type = std::tuple<Args...>;
    };

// C5046: Symbol involving type with internal linkage not defined
#pragma warning(disable: 5046)

    template<typename F>
    auto arguments(const F&) -> typename signature<F>::type;

    void throwIfRequired(HRESULT hr)
    {
      if (FAILED(hr))
      {
        auto error = _com_error(hr);
        XLO_THROW(L"COM Error {0:#x}: {1}", (unsigned)error.Error(), error.ErrorMessage());
      }
    }

    template<class T>
    _variant_t toVariant(const T& x)
    {
      return _variant_t(x);
    }
    
    template<>
    _variant_t toVariant(const std::wstring_view& str)
    {
      auto variant = COM::stringToVariant(str);
      return _variant_t(variant, false);
    }

    template<class TObj, class V = TObj::get_Item>
    IDispatch* getItemHelper(TObj& obj, const _variant_t& what, HRESULT& retCode)
    {
      auto resultType = std::get<1>(arguments(&TObj::get_Item));
      std::remove_pointer_t<decltype(resultType)> result;
      retCode = obj.get_Item(what, &result);
      return result;
    }

    template<class TObj>
    IDispatch* getItemHelper(TObj& obj, const _variant_t& what, HRESULT& retCode)
    {
      try
      {
        retCode = S_OK;
        return obj.GetItem(what).Detach();
      }
      catch (_com_error& error)
      {
        retCode = error.Error();
        return nullptr;
      }
    }

    template<class TObj, class TRes, class TArg>
    bool comTryGet(TObj& obj, const TArg& what, TRes& out)
    {
      HRESULT hr;
      auto result = getItemHelper(obj, toVariant(what), hr);
      if (hr == DISP_E_BADINDEX)
        return false;

      throwIfRequired(hr);

      out = TRes((ComType<TRes>*)result, true);
      return true;
    }

    template<class TRes, class TObj, class TArg>
    TRes comGetItem(TObj& obj, const TArg& what)
    {
      TRes result;
      auto found = comTryGet(obj, what, result);
      if (!found)
        XLO_THROW(L"Collection: could not find '{}'", what);
      return std::move(result);
    }

    template<class T, class V=T::get_Application>
    auto comGetApp(T& x)
    {
      struct Excel::_Application* _result = 0;
      HRESULT hr = x.get_Application(&_result);
      throwIfRequired(hr);
      return Application(_result);
    }

    template<class T>
    auto comGetApp(T& x)
    {
      try
      {
        return Application(x.Application.Detach());
      }
      XLO_RETHROW_COM_ERROR;
    }

    template<class T, class V = T::get_Count>
    auto comGetCount(T& x)
    {
      long result = 0;
      HRESULT hr = com().get_Count(&result);
      throwIfRequired(hr);
      return result
    }

    template<class T>
    auto comGetCount(T& x)
    {
      try
      {
        return x.GetCount();
      }
      XLO_RETHROW_COM_ERROR;
    }
  }

  Application& thisApp()
  {
    return COM::attachedApplication();
  }

  void detail::UnknownObject::release()
  {
    if (_ptr)
    {
      _ptr->Release();
      _ptr = nullptr;
    }
  }

  void detail::UnknownObject::init(IUnknown* ptr, bool steal)
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
    Application workbookFinder(const wchar_t* workbook)
    {
      HWND xlmain = 0;
      while ((xlmain = COM::nextExcelMainWindow(xlmain)) != 0)
      {
        auto xlApp = Application(size_t(xlmain));
        ExcelWorkbook wb(nullptr);
        if (xlApp.workbooks().tryGet(workbook, wb))
          return xlApp;
      }
      throw ComConnectException("Failed to create Application object from workbook");
    }
  }

  Application::Application(const wchar_t* workbook)
    : Application(workbookFinder(workbook))
  {}

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

  ExcelWorksheet Application::activeWorksheet() const
  {
    try
    {
      Excel::_Worksheet* sheet = nullptr;
      com().ActiveSheet->QueryInterface(&sheet);
      return ExcelWorksheet(sheet);
    }
    XLO_RETHROW_COM_ERROR;
  }

  ExcelRange Application::activeCell() const
  {
    try
    {
      Excel::Range* range = nullptr;
      com().ActiveCell->QueryInterface(&range);
      return ExcelRange(range);
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

  bool Application::getEnableEvents() const
  {
    try
    {
      return com().EnableEvents == VARIANT_TRUE;
    }
    XLO_RETHROW_COM_ERROR;
  }

  bool Application::setEnableEvents(bool value)
  {
    try
    {
      auto& obj = com();
      auto previousValue = obj.EnableEvents == VARIANT_TRUE;
      obj.EnableEvents = _variant_t(value);
      return previousValue;
    }
    XLO_RETHROW_COM_ERROR;
  }

  bool Application::getDisplayAlerts() const
  {
    try
    {
      return com().GetDisplayAlerts() == VARIANT_TRUE;
    }
    XLO_RETHROW_COM_ERROR;
  }

  bool Application::setDisplayAlerts(bool value)
  {
    try
    {
      auto& obj = com();
      auto previousValue = obj.GetDisplayAlerts() == VARIANT_TRUE;
      obj.PutDisplayAlerts(0, value ? VARIANT_TRUE : VARIANT_FALSE);
      return previousValue;
    }
    XLO_RETHROW_COM_ERROR;
  }

  bool Application::getScreenUpdating() const
  {
    try
    {
      return com().GetScreenUpdating() == VARIANT_TRUE;
    }
    XLO_RETHROW_COM_ERROR;
  }

  bool Application::setScreenUpdating(bool value)
  {
    try
    {
      auto& obj = com();
      auto previousValue = obj.GetScreenUpdating() == VARIANT_TRUE;
      obj.PutScreenUpdating(0, value ? VARIANT_TRUE : VARIANT_FALSE);
      return previousValue;
    }
    XLO_RETHROW_COM_ERROR;
  }

  Application::CalculationMode Application::getCalculationMode() const
  {
    try
    {
      return (CalculationMode)com().GetCalculation();
    }
    XLO_RETHROW_COM_ERROR;
  }

  Application::CalculationMode Application::setCalculationMode(CalculationMode value)
  {
    try
    {
      auto& obj = com();
      auto previousValue = obj.GetCalculation();
      obj.PutCalculation(0, (Excel::XlCalculation)value);
      return (CalculationMode)previousValue;
    }
    XLO_RETHROW_COM_ERROR;
  }

  ExcelRange Application::selection()
  {
    try
    {
      return fromComPtr<ExcelRange>(com().Selection);
    }
    XLO_RETHROW_COM_ERROR;
  }

  namespace
  {
    template <typename F, typename T, std::size_t N, std::size_t... Idx>
    decltype(auto) appRun_impl(F func, T(&args)[N], std::index_sequence<Idx...>) {
      return thisApp().com().Run(func, args[Idx]...);
    }

    template <typename T, std::size_t N>
    decltype(auto) appRun(const wchar_t* func, T(&args)[N]) {
      return appRun_impl(func, args, std::make_index_sequence<N>{});
    }
  }

  ExcelObj Application::run(
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

  ExcelWorkbook Application::open(
    const std::wstring& filepath, 
    bool updateLinks, 
    bool readOnly,
    wchar_t delimiter)
  {
    try
    {
      return fromComPtr<ExcelWorkbook>(com().Workbooks->Open(
        _bstr_t(filepath.c_str()),
        updateLinks ? 3 : 0,
        _variant_t(readOnly),
        delimiter == 0 ? 5 : 6,
        vtMissing,
        vtMissing,
        vtMissing,
        vtMissing,
        delimiter != 0 ? _variant_t(wstring(delimiter, 1).c_str()) : vtMissing
      ));
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
            return app.com().Windows->GetItem(toVariant(caption)).Detach();
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
    try
    {
      return Application(com().Application.Detach());
    }
    XLO_RETHROW_COM_ERROR;
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
            return workbooks->GetItem(toVariant(name)).Detach();
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
    return comGetApp(com());
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

      auto ws = fromComPtr<ExcelWorksheet>(com().Worksheets->Add(
        before.valid() ? _variant_t(&before.com()) : vtMissing,
        after.valid()  ? _variant_t(&after.com())  : vtMissing));
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
        com().SaveAs(toVariant(filepath), 
          vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, 
          Excel::XlSaveAsAccessMode::xlNoChange);
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
    return comGetApp(com());
  }

  ExcelWorkbook ExcelWorksheet::parent() const
  {
    try
    {
      return fromComPtr<ExcelWorkbook>(com().Parent);
    }
    XLO_RETHROW_COM_ERROR;
  }

  ExcelRange ExcelWorksheet::range(
    int fromRow, int fromCol,
    int toRow, int toCol) const
  {
    try
    {
      auto& obj = com();
      if (toRow == Range::TO_END)
        toRow = obj.Rows->GetCount();
      if (toCol == Range::TO_END)
        toCol = obj.Columns->GetCount();

      auto r = obj.GetRange(
        obj.Cells->Item[fromRow + 1][fromCol + 1],
        obj.Cells->Item[toRow + 1][toCol + 1]);
      return ExcelRange(r);
    }
    XLO_RETHROW_COM_ERROR;
  }

  ExcelRange ExcelWorksheet::range(const std::wstring_view& address) const
  {
    try
    {
      return ExcelRange(
        formatStr(L"'[%s]%s'!%s", 
          parent().name().c_str(), 
          name().c_str(), 
          wstring(address).data()),
        app());
    }
    XLO_RETHROW_COM_ERROR;
  }

  ExcelObj ExcelWorksheet::value(Range::row_t i, Range::col_t j) const
  {
    Excel::RangePtr pRange(com().Cells->Item[i][j].pdispVal);
    return COM::variantToExcelObj(pRange->Value2);
  }

  ExcelRange ExcelWorksheet::usedRange() const
  {
    try
    {
      return ExcelRange(com().GetUsedRange(0));
    }
    XLO_RETHROW_COM_ERROR;
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
      com().Name = toVariant(name).bstrVal;
    }
    XLO_RETHROW_COM_ERROR;
  }

  ExcelWorkbook Workbooks::add()
  {
    try
    {
      return fromComPtr<ExcelWorkbook>(com().Add());
    }
    XLO_RETHROW_COM_ERROR;
  }

  namespace
  {
    auto variantToUnknown(const VARIANT& v)
    {
      if (v.vt != VT_DISPATCH)
        XLO_THROW("Unexpected variant type, should be IDispatch");
      return UnknownObject(v.pdispVal, true);
    }
  }

  detail::ComIteratorBase::ComIteratorBase(
    IUnknown* ptr,
    UnknownObject next)
    : AppObject([=]() {
        IEnumVARIANT* iterator;
        ptr->QueryInterface(&iterator);
        return iterator;
      }(), true)
    , _next(next)
  {
  }

  UnknownObject detail::ComIteratorBase::get()
  {
    return _next;
  }

  void detail::ComIteratorBase::increment()
  {
    try
    {
      VARIANT next;
      ULONG nFetched;
      if (SUCCEEDED(com().Next(1, &next, &nFetched)) && nFetched > 0)
        _next = variantToUnknown(next);
      else
        _next = UnknownObject();
    }
    XLO_RETHROW_COM_ERROR;
  }

  detail::ComIteratorBase detail::ComIteratorBase::excrement()
  {
    try
    {
      IEnumVARIANT* clone;
      if (FAILED(com().Clone(&clone)))
        XLO_THROW("ComIterator: internal error copy failed");

      auto prev = UnknownObject(std::move(_next));
      increment();

      return ComIteratorBase(clone, prev);
    }
    XLO_RETHROW_COM_ERROR;
  }

  bool detail::ComIteratorBase::operator==(const detail::ComIteratorBase& other) const
  {
    return _next.ptr() == other._next.ptr();
  }

  void detail::ComIteratorBase::getMany(size_t n, std::vector<UnknownObject>& result)
  {
    try
    {
      // We take the value in _next and fetch *n* more values 
      vector<VARIANT> variants(n);
      ULONG nFetched;
      com().Next((ULONG)n, variants.data(), &nFetched);
    
      result.emplace_back(std::move(_next));

      for (auto i = 0u; i < std::min(n - 1, (size_t)nFetched); ++i)
        result.emplace_back(variantToUnknown(variants[i]));
      
      // If iterator was exhausted, leave _next as null value, otherwise
      // take the last value we fetched
      _next = variantToUnknown(variants[n - 1]);
    }
    XLO_RETHROW_COM_ERROR;
  }

  template<class T, class Ptr>
  T Collection<T, Ptr>::get(const std::wstring_view& name) const
  {
    return comGetItem<T>(com(), name);
  }

  template<class T, class Ptr>
  T Collection<T, Ptr>::get(const size_t index) const
  {
    return comGetItem<T>(com(), index);
  }

  template<class T, class Ptr>
  bool Collection<T, Ptr>::tryGet(const std::wstring_view& name, T& out) const
  {
    return comTryGet(com(), name, out);
  }

  template<class T, class Ptr>
  bool Collection<T, Ptr>::tryGet(const size_t index, T& out) const
  {
    return comTryGet(com(), index, out);
  }

  template<class T, class Ptr>
  size_t Collection<T, Ptr>::count() const
  {
    return comGetCount(com());
  }

  template<class T, class Ptr>
  Application Collection<T, Ptr>::app() const
  {
    return comGetApp(com());
  }

  template<class T, class Ptr>
  std::vector<T> Collection<T, Ptr>::list() const
  {
    try
    {
      auto iterator = ComIterator<T>((IUnknown*)com().Get_NewEnum());
      return iterator.getMany(count());
    }
    XLO_RETHROW_COM_ERROR;
  }
  
  template<class T, class Ptr>
  ComIterator<T> Collection<T, Ptr>::begin() const
  {
    return ComIterator<T>(com().Get_NewEnum());
  }

  Worksheets::Worksheets(const Application& app)
    : Worksheets(app.workbooks().active())
  {}

  Worksheets::Worksheets(const ExcelWorkbook& workbook)
    : Collection(workbook.valid()
      ? workbook.com().Worksheets.Detach()
      : nullptr)
  {
    if (!valid())
      XLO_THROW("No active workbook or workbook invalid");
  }

  ExcelWorkbook Worksheets::parent() const
  {
    return fromComPtr<ExcelWorkbook>(com().Parent);
  }

  Workbooks::Workbooks(const Application& app)
    : Collection(app.com().Workbooks.Detach())
  {}

  Windows::Windows(const Application& app)
    : Collection(app.com().Windows.Detach())
  {}

  Windows::Windows(const ExcelWorkbook& workbook)
    : Collection(workbook.com().Windows.Detach())
  {}

  //template<>
  //bool Collection<ExcelWindow, Excel::Windows>::tryGet(
  //  const std::wstring_view& name, ExcelWindow& out) const
  //{
  //  try
  //  {
  //    out = fromComPtr<ExcelWindow>(com().GetItem(toVariant(name)));
  //    return true;
  //  }
  //  catch (_com_error& error)
  //  {
  //    if (error.Error() == DISP_E_BADINDEX)
  //      return false;
  //    XLO_THROW(L"COM Error {0:#x}: {1}", (size_t)error.Error(), error.ErrorMessage());
  //  }
  //}

  Ranges::Ranges(const ExcelRange& multiRange)
    : Collection(multiRange.com().Areas.Detach())
  { }

  const std::set<std::wstring>& Application::workbookPaths()
  {
    return COM::workbookPaths();
  }
}

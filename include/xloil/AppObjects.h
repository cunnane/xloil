#pragma once
#include <xloil/ExportMacro.h>
#include <xloil/ExcelObj.h>
#include <xloil/Range.h>
#include <xloil/ExcelRef.h>
#include <string>
#include <memory>
#include <vector>
#include <set>

// Forward Declarations from Typelib
struct IDispatch;
struct IUnknown;
struct IEnumUnknown;
struct IEnumVARIANT;

namespace Excel 
{
  struct _Application;
  struct Window;
  struct _Workbook;
  struct _Worksheet;
  struct Range;
  struct Windows;
  struct Workbooks;
  struct Sheets;
  struct Areas;
}

namespace xloil
{
  class ExcelWindow;
  class ExcelWorksheet;
  class Windows;
  class Workbooks;
  class Worksheets;
  class ExcelWorkbook;
  class ExcelRange;
  class Application;
  class Ranges;
}

namespace xloil
{
  class ComConnectException : public std::runtime_error
  {
  public:
    ComConnectException(const char* message)
      : std::runtime_error(message)
    {}
  };

  /// <summary>
  /// Thrown if an AppObject has a null underlying pointer.  Currently
  /// only the Application object checks this
  /// </summary>
  class NullComObjectException : public std::exception
  {
  public:
    NullComObjectException(const char* message)
      : std::exception(message)
    {}
    NullComObjectException()
      : std::exception()
    {}
  };

  namespace detail
  {
    /// <summary>
    /// IUnknown ptr holder. Used internally.
    /// </summary>
    class UnknownObject
    {
    public:
      UnknownObject(IUnknown* ptr = nullptr, bool steal = false)
      {
        init(ptr, steal);
      }

      UnknownObject(UnknownObject&& that) noexcept
        : _ptr(nullptr)
      {
        std::swap(_ptr, that._ptr);
      }

      UnknownObject(const UnknownObject& that) 
        : UnknownObject(that._ptr, false) 
      {}

      UnknownObject& operator=(const UnknownObject& that) noexcept
      {
        release();
        init(that._ptr);
        return *this;
      }

      UnknownObject& operator=(UnknownObject&& that) noexcept
      {
        release();
        std::swap(_ptr, that._ptr);
        return *this;
      }

      ~UnknownObject()
      {
        release();
      }

      IUnknown* ptr() const { return _ptr; }
      bool valid() const { return _ptr; }
      XLOIL_EXPORT void release();

    private:
      IUnknown* _ptr;

      XLOIL_EXPORT void init(IUnknown* ptr, bool steal = false);
    };

    template <typename T,
#ifdef NDEBUG
      bool TCheck = false>
#else
      bool TCheck = true >
#endif
    class AppObject : public UnknownObject
    {
    public:
      AppObject(T* ptr = nullptr, bool steal = false)
        : UnknownObject((IUnknown*)(ptr), steal)
      {}

      AppObject(const UnknownObject& obj)
        : UnknownObject(obj)
      {}

      AppObject(UnknownObject&& obj)
        : UnknownObject(obj)
      {}

      void check() const
      {
        if constexpr (TCheck)
        {
          if (!valid())
            throw new NullComObjectException();
        }
      }

      T& com() const { check(); return *(T*)ptr(); }
    };
 
    // C4661: no suitable definition provided for explicit template instantiation request
    #pragma warning(disable: 4661)

    class XLOIL_EXPORT ComIteratorBase : public AppObject<IEnumVARIANT>
    {
    public:
      ComIteratorBase(IUnknown* ptr, UnknownObject next);
      ComIteratorBase(IUnknown* ptr)
        : ComIteratorBase(ptr, UnknownObject())
      {
        increment();
      }
      ComIteratorBase()
        : AppObject(nullptr)
        , _next(nullptr)
      {}

      UnknownObject get();
      void increment();
      ComIteratorBase excrement();
      bool operator==(const ComIteratorBase& other) const;

      void getMany(size_t n, std::vector<UnknownObject>& result);

    private:
      UnknownObject _next;
    };
  }

  template<class T>
  class ComIterator : private detail::ComIteratorBase
  {
  public:
    using detail::ComIteratorBase::ComIteratorBase;

    T operator*()
    {
      return T(get());
    }
    ComIterator& operator++()
    {
      increment();
      return *this;
    }
    ComIterator operator++(int)
    {
      return (ComIterator)excrement();
    }
    std::vector<T> getMany(size_t n)
    {
      std::vector<UnknownObject> objects;
      std::vector<T> result;
      detail::ComIteratorBase::getMany(n, objects);
      std::transform(objects.cbegin(), objects.cend(), std::back_inserter(result),
        [](auto x) { return T(x); });
      return result;
    }
    bool operator==(const ComIterator<T>& that) const
    {
      return detail::ComIteratorBase::operator==(that);
    }
  };


  template<class T, class Ptr>
  class Collection : public detail::AppObject<Ptr>
  {
  public:
    Collection(Ptr* ptr)
      : AppObject(ptr, true)
    {}

    T get(const std::wstring_view& name) const;
    T get(const size_t index) const;
    bool tryGet(const std::wstring_view& name, T& wb) const;
    bool tryGet(const size_t index, T& wb) const;

    auto operator[](const std::wstring_view& name) const
    {
      return get(name);
    };

    std::vector<T> list() const;
    size_t count() const;

    Application app() const;

    ComIterator<T> begin() const;
    ComIterator<T> end() const
    {
      return ComIterator<T>();
    }
  };


  class XLOIL_EXPORT Application : public detail::AppObject<Excel::_Application, true>
  {
  public:
    /// <summary>
    /// Construct an application object from a ptr to the underlying COM object.
    /// If no ptr is provided, a new instance of Excel is created. Always steals
    /// a reference, i.e. does not call AddRef.
    /// </summary>
    Application(Excel::_Application* app = nullptr);
    /// <summary>
    /// Construct an application object from a window handle to the top level
    /// Excel window (which has window class XLMAIN)
    /// </summary>
    Application(size_t hWnd);
    /// <summary>
    /// Construct an application object from a workbook name. This searches all
    /// open Excel instances for the one which has opened the workbook. It will
    /// fail if the workbook is not open
    /// </summary>
    Application(const wchar_t* workbook);

    std::wstring name() const;

    /// <summary>
    /// Calculates
    /// </summary>
    void calculate(const bool full=false, const bool rebuild=false);

    Workbooks workbooks() const;
    Windows windows() const;
    ExcelWorksheet activeWorksheet() const;
    ExcelRange activeCell() const;

    ExcelObj run(const std::wstring& func, const size_t nArgs, const ExcelObj* args[]);

    ExcelWorkbook open(
      const std::wstring& filepath, 
      bool updateLinks=true, 
      bool readOnly=false,
      wchar_t delimiter = 0);

    /// <summary>
    /// The set of full path names of all open workbooks
    /// </summary>
    const std::set<std::wstring>& workbookPaths();

    /// <summary>
    /// Calls Application.Quit to close the Excel instance and frees the COM resources.
    /// This invalidates the Application object: any further calls to methods other 
    /// than quit() will raise an exception.
    /// </summary>
    /// <param name="silent">
    ///   If true, supresses save file dialogs: unsaved changes to workbooks will be discarded.
    /// </param>
    void quit(bool silent=true);

    bool getVisible() const;
    void setVisible(bool x);

    bool getEnableEvents() const;
    bool setEnableEvents(bool value);

    bool getDisplayAlerts() const;
    bool setDisplayAlerts(bool value);

    bool getScreenUpdating() const;
    bool setScreenUpdating(bool value);

    enum CalculationMode
    {
      Automatic = -4105,
      Manual = -4135,
      Semiautomatic = 2
    };

    CalculationMode Application::getCalculationMode() const;
    CalculationMode Application::setCalculationMode(CalculationMode value);

    /// <summary>
    /// Returns an invalid ExcelRange is the selection is not a range
    /// </summary>
    ExcelRange selection();
  };

  /// <summary>
  /// Gets the Excel.Application object which is the root of the COM API 
  /// </summary>
  XLOIL_EXPORT Application& thisApp();

  enum class SpecialCells : int
  {
    Blanks = 4,
    Constants = 2,
    Formulas = -4123,
    LastCell = 11,
    Comments = -4144,
    Visible = 12,
    AllFormatConditions = -4172,
    SameFormatConditions = -4173,
    AllValidation = -4174,
    SameValidation = -4175
  };

  class XLOIL_EXPORT ExcelRange : public Range, public detail::AppObject<Excel::Range>
  {
  public:
    /// <summary>
    /// Constructs a Range from a address. A local address (not qualified with
    /// a workbook name) will refer to the active workbook
    /// </summary>
    explicit ExcelRange(
      const std::wstring_view& address,
      const Application& app = thisApp());

    ExcelRange(const Range& range);
    ExcelRange(const ExcelRef& ref) : ExcelRange(ref.address()) {}

    using detail::AppObject<Excel::Range>::AppObject;

    std::unique_ptr<Range> range(
      int fromRow, int fromCol,
      int toRow = TO_END, int toCol = TO_END) const final override;

    std::unique_ptr<Range> trim() const final override;

    std::tuple<row_t, col_t> shape() const final override;

    std::tuple<row_t, col_t, row_t, col_t> bounds() const final override;

    std::wstring address(AddressStyle style = AddressStyle::A1) const final override;

    size_t nAreas() const;

    ExcelObj value() const final override;

    ExcelObj value(row_t i, col_t j) const final override;

    void set(const ExcelObj& value) final override;

    ExcelObj formula() const final override;

    std::optional<bool> hasFormula() const final override;

    void clear() final override;

    virtual Excel::Range* asComPtr() const final override
    {
      return &com();
    }

    enum SetFormulaMode
    {
      ARRAY_FORMULA,
      DYNAMIC_ARRAY,
      OLD_ARRAY
    };

    /// <summary>
    /// Sets the forumula for the range to the specified string. The `mode` 
    /// parameter determines how this function differs from the *Formula2* 
    /// property of COM/VBA Range:
    ///
    ///   * *DYNAMIC_ARRAY*: (default) identical the `Formula2` property, formulae
    ///    which return arrays will spill.  If the range is larger than one cell and 
    ///    a single value is passed that value is filled into each cell.
    ///   * *ARRAY_FORMULA*: if the target range is larger than one cell and a single 
    ///    string is passed, the string is set as an array formula for the range
    ///   * *OLD_ARRAY*: formulae which return arrays will not spill see "Formula vs Formula2" 
    ///    on MSDN
    /// 
    /// </summary>
    /// <param name="formula"></param>
    /// 
    void setFormula(const std::wstring_view& formula, const SetFormulaMode mode = DYNAMIC_ARRAY);
    
    /// <summary>
    /// Instead of taking only a string formula, takes an *ExcelObj* which can contain a string
    /// or an array of equal dimensions to the *Range* being set.
    /// </summary>
    void setFormula(const ExcelObj& formula, const SetFormulaMode mode = DYNAMIC_ARRAY);

    /// <summary>
    /// The range address
    /// </summary>
    std::wstring name() const;

    /// <summary>
    /// The worksheet which contains this range
    /// </summary>
    ExcelWorksheet parent() const;

    /// <summary>
    /// Returns the Application object which owns this Range
    /// </summary>
    Application app() const;

    ExcelRange specialCells(SpecialCells type, 
                            ExcelType values = ExcelType(0)) const;

    Ranges areas() const;


    ComIterator<ExcelRange> begin() const;
    ComIterator<ExcelRange> end() const
    {
      return ComIterator<ExcelRange>();
    }
  };


  /// <summary>
  /// Wraps an COM Excel::Window object to avoid exposing the COM typelib
  /// </summary>
  class XLOIL_EXPORT ExcelWorksheet : public detail::AppObject<Excel::_Worksheet>
  {
  public:
    using detail::AppObject<Excel::_Worksheet>::AppObject;

    /// <summary>
    /// Returns the window title
    /// </summary>
    /// <returns></returns>
    std::wstring name() const;

    /// <summary>
    /// Gives the name of the workbook which owns this sheet
    /// </summary>
    ExcelWorkbook parent() const;

    /// <summary>
    /// Returns the underlying Application object, the ultimate parent 
    /// of this object
    /// </summary>
    Application app() const;

    /// <summary>
    /// Returns a range on this worksheet given a (local) address
    /// </summary>
    ExcelRange range(const std::wstring_view& address) const;

    /// <summary>
    /// Returns the range on this worksheet starting and ending at the specified
    /// rows and columns.  Includes the start row and columns and the end row and 
    /// column, just as Excel's Worksheet.Range behaves.
    /// </summary>
    ExcelRange range(int fromRow, int fromCol,
      int toRow = ExcelRange::TO_END, int toCol = ExcelRange::TO_END) const;

    /// <summary>
    /// Returns a Range referring to the single cell (i, j) (zero-based indexing)
    /// </summary>
    ExcelRange cell(int i, int j) const
    {
      return range(i, j, i, j);
    }

    /// <summary>
    /// Convenience wrapper for cell(i,j)->value()
    /// </summary>
    ExcelObj value(Range::row_t i, Range::col_t j) const;

    /// <summary>
    /// Returns a ExcelRange object that represents the used range on the worksheet.
    /// </summary>
    ExcelRange usedRange() const;

    /// <summary>
    /// Returns the size of the worksheet, which is always (MaxRows, MaxCols).
    /// This function exists mainly to provide some polymorphism with Range.
    /// </summary>
    std::tuple<Range::row_t, Range::col_t> shape() const
    {
      return std::make_tuple(XL_MAX_ROWS, XL_MAX_COLS);
    }

    /// <summary>
    /// Makes this worksheet the active one in its workbook
    /// </summary>
    void activate();

    /// <summary>
    /// Calculates this worksheet
    /// </summary>
    void calculate();

    /// <summary>
    /// Sets the worksheet name (note 31 char limit)
    /// </summary>
    void setName(const std::wstring_view& name);
  };


  /// <summary>
  /// Wraps a Workbook (https://docs.microsoft.com/en-us/office/vba/api/excel.workbook) in
  /// Excel's object model but with very limited functionality at present
  /// </summary>
  class XLOIL_EXPORT ExcelWorkbook : public detail::AppObject<Excel::_Workbook>
  {
  public:
    /// <summary>
    /// Gives the ExcelWorkbook object associated with the given workbookb name, or the active workbook
    /// </summary>
    /// <param name="name">The name of the workbook to find, or the active workbook if null</param>
    explicit ExcelWorkbook(
      const std::wstring_view& name = std::wstring_view(), 
      Application app = thisApp());
    
    using detail::AppObject<Excel::_Workbook>::AppObject;

    std::wstring name() const;

    /// <summary>
    /// Returns the full file path and file name for this workbook
    /// </summary>
    std::wstring path() const;

    /// <summary>
    /// Returns the underlying Application object, the ultimate parent 
    /// of this object
    /// </summary>
    Application app() const;

    /// <summary>
    /// Returns a list of all Windows associated with this workbook (there can be
    /// multiple windows viewing a single workbook).
    /// </summary>
    /// <returns></returns>
    Windows windows() const;

    /// <summary>
    /// Returns a collection of all Worksheets in this Workbook. It does not  
    /// include chart sheets (or Excel4 macro sheets)
    /// </summary>
    /// <returns></returns>
    Worksheets worksheets() const;

    /// <summary>
    /// Returns a worksheet object corresponding to the specified name if 
    /// it exists
    /// </summary>
    ExcelWorksheet worksheet(const std::wstring_view& name) const;

    /// <summary>
    /// Returns a range in this workbook given an address
    /// </summary>
    ExcelRange range(const std::wstring_view& address) const
    {
      return ExcelRange(address, app());
    }

    /// <summary>
    /// Adds a new worksheet, naming it if a name is provided, otherwise it
    /// will have a default name provided by Excel, such as 'Sheet4'.
    /// </summary>
    /// <param name="name"></param>
    /// <param name="before"></param>
    /// <param name="after"></param>
    /// <returns></returns>
    ExcelWorksheet add(
        const std::wstring_view& name = std::wstring_view(),
        const ExcelWorksheet& before = ExcelWorksheet(nullptr),
        const ExcelWorksheet& after = ExcelWorksheet(nullptr)) const;

    /// <summary>
    /// Makes this the active workbook
    /// </summary>
    /// <returns></returns>
    void activate() const;

    void save(const std::wstring_view& filepath = std::wstring_view());

    void close(bool save=true);
  };


  /// <summary>
  /// Wraps an COM Excel::Window object to avoid exposing the COM typelib
  /// </summary>
  class XLOIL_EXPORT ExcelWindow : public detail::AppObject<Excel::Window>
  {
  public:
    using detail::AppObject<Excel::Window>::AppObject;

    /// <summary>
    /// Gives the ExcelWindow object associated with the given window caption, or the active window
    /// </summary>
    /// <param name="caption">The name of the window to find, or the active window if null</param>
    explicit ExcelWindow(
      const std::wstring_view& caption = std::wstring_view(),
      Application app = thisApp());

    /// <summary>
    /// Retuns the Win32 window handle
    /// </summary>
    /// <returns></returns>
    size_t hwnd() const;

    /// <summary>
    /// Returns the window title
    /// </summary>
    std::wstring name() const;

    /// <summary>
    /// Returns the underlying Application object, the ultimate parent 
    /// of this object
    /// </summary>
    Application app() const;

    /// <summary>
    /// Gives the name of the workbook displayed by this window 
    /// </summary>
    ExcelWorkbook workbook() const;
  };

  inline std::wstring to_wstring(const ExcelRange& x) { return x.name(); }

  XLOIL_EXPORT ExcelRef refFromComRange(Excel::Range& range);

  inline ExcelRef refFromRange(const Range& range)
  {
    auto excelRange = dynamic_cast<const ExcelRange*>(&range);
    if (excelRange)
      return refFromComRange(excelRange->com());
    else
      return static_cast<const XllRange&>(range).asRef();
  }


  class XLOIL_EXPORT Worksheets : public Collection<ExcelWorksheet, Excel::Sheets>
  {
  public:
    Worksheets(const Application& app = thisApp());
    Worksheets(const ExcelWorkbook& workbook);
    ExcelWorksheet active() const { return app().activeWorksheet(); }
    ExcelWorksheet add(
      const std::wstring_view& name = std::wstring_view(),
      const ExcelWorksheet& before = ExcelWorksheet(nullptr),
      const ExcelWorksheet& after = ExcelWorksheet(nullptr)) const
    {
      return parent().add(name, before, after);
    }
    ExcelWorkbook parent() const;
  };


  class XLOIL_EXPORT Workbooks : public Collection<ExcelWorkbook, Excel::Workbooks>
  {
  public:
    Workbooks(const Application& app = thisApp());
    ExcelWorkbook active() const 
    {
      return ExcelWorkbook(std::wstring_view(), app());
    }
    ExcelWorkbook add();
  };


  class XLOIL_EXPORT Windows : public Collection<ExcelWorksheet, Excel::Windows>
  {
  public:
    Windows(const Application& app = thisApp());
    Windows(const ExcelWorkbook& workbook);
    ExcelWindow active() const
    {
      return ExcelWindow(std::wstring_view(), app());
    }
  };


  class XLOIL_EXPORT Ranges : public Collection<ExcelRange, Excel::Areas>
  {
  public:
    Ranges(const ExcelRange& multiRange);
  };


  class PauseExcel
  {
  private:
    Application _app;
    Application::CalculationMode _previousCalculation;
    bool _previousEvents;
    bool _previousAlerts;
    bool _previousUpdating;

  public:
    PauseExcel(Application& app)
      : _app(app)
      , _previousCalculation(app.setCalculationMode(Application::Manual))
      , _previousAlerts(app.setDisplayAlerts(false))
      , _previousEvents(app.setEnableEvents(false))
      , _previousUpdating(app.setScreenUpdating(false))
    {}
    ~PauseExcel()
    {
      _app.setCalculationMode(_previousCalculation);
      _app.setDisplayAlerts(_previousAlerts);
      _app.setEnableEvents(_previousEvents);
      _app.setScreenUpdating(_previousUpdating);
    }
  };

  // Some function definitions which need to live down here due to
  // the order of declarations

  inline Workbooks Application::workbooks() const
  {
    return Workbooks(*this);
  }

  inline Windows Application::windows() const
  {
    return Windows(*this);
  }

  inline Worksheets ExcelWorkbook::worksheets() const
  { 
    return Worksheets(*this); 
  }

  inline Windows ExcelWorkbook::windows() const
  {
    return Windows(*this);
  }

  inline ExcelWorksheet ExcelWorkbook::worksheet(const std::wstring_view& name) const
  {
    return worksheets().get(name);
  }

  inline Ranges ExcelRange::areas() const
  {
    return Ranges(*this);
  }
}

template<class T>
struct std::iterator_traits<xloil::ComIterator<T>>
{
  using iterator_category = std::forward_iterator_tag;
  using value_type = T;
  using reference = const T&;
  using pointer = const T*;
  using difference_type = size_t;
};
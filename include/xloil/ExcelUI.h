#pragma once
#include <xloil/ExportMacro.h>
#include <xloil/ExcelObj.h>
#include <functional>
#include <memory>
#include <map>

typedef struct tagVARIANT VARIANT;
struct IDispatch;

namespace Excel { struct Window; }
namespace xloil { class ExcelWindow; }
namespace xloil
{
  /// <summary>
  /// This class is passed to Ribbon callback functions and identifies which
  /// control triggered the callback.
  /// </summary>
  struct RibbonControl
  {
    const wchar_t* Id;
    const wchar_t* Tag;
  };

  /// <summary>
  /// Displays a message in Excel's status bar
  /// </summary>
  /// <param name="msg"></param>
  /// <param name="timeout">if positive, the message will be cleared after the specified
  /// number of milliseconds</param>
  /// <returns></returns>
  XLOIL_EXPORT void statusBarMsg(const std::wstring_view& msg, size_t timeout = 0);

  /// <summary>
  /// Displays status bar messages, then clears the status bar on destruction
  /// (with an optional delay)
  /// </summary>
  class StatusBar
  {
  public:
    /// <summary>
    /// Sets a final timeout in milliseconds. The status bar will be cleared
    /// after this amount of time since the class is destroyed.
    /// </summary>
    /// <param name="finalTimeout"></param>
    StatusBar(size_t finalTimeout = 0) 
      : _timeout(finalTimeout) 
    {}

    void msg(const std::wstring_view& msg, size_t timeout = 0)
    {
      statusBarMsg(msg, timeout);
    }

    ~StatusBar()
    {
      statusBarMsg(L"", _timeout);
    }

    size_t _timeout;
  };

  /// <summary>
  /// Event handler to respond to custom task pane events
  /// </summary>
  class ICustomTaskPaneHandler
  {
  public:
    virtual void onSize(int width, int height) = 0;
    virtual void onVisible(bool c) = 0;
    virtual void onDocked() = 0;
    virtual void onDestroy() = 0;
  };

  /// <summary>
  /// Wraps an Excel Custom Task Pane 
  /// (https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.tools.customtaskpane)
  /// </summary>
  class ICustomTaskPane
  {
  public:
    enum DockPosition
    {
      Bottom =	3,
      Floating =	4,
      Left =	0,
      Right = 2,
      Top = 1
    };
    virtual ~ICustomTaskPane() {}

    /// <summary>
    /// Gets the base COM control created in the task pane. Thread safe.
    /// </summary>
    /// <returns></returns>
    virtual IDispatch* content() const = 0;

    /// <summary>
    /// Gives the main Excel Window object to which this task pane is attached
    /// </summary>
    /// <returns></returns>
    virtual ExcelWindow window() const = 0;

    /// <summary>
    /// Gets the window handle for parent of the task pane control, that is
    /// the child window which houses the task pane. Thread safe.
    /// </summary>
    /// <returns></returns>
    virtual size_t parentWindowHandle() const = 0;

    /// <summary>
    /// Sets the visblity of the task pane to true/false
    /// </summary>
    /// <param name="visible"></param>
    virtual void setVisible(bool visible) = 0;

    /// <summary>
    /// Returns the current visibility of the task pane
    /// </summary>
    /// <returns></returns>
    virtual bool getVisible() = 0;

    /// <summary>
    /// Returns the task pane size in pixels as a tuple (width, height)
    /// </summary>
    /// <returns></returns>
    virtual std::pair<int, int> getSize() = 0;

    /// <summary>
    /// Sets the task pane size to the specified number of pixels
    /// </summary>
    /// <param name="width"></param>
    /// <param name="height"></param>
    virtual void setSize(int width, int height) = 0;

    /// <summary>
    /// Gets the current dock position of the task pane
    /// </summary>
    /// <returns></returns>
    virtual DockPosition getPosition() const = 0;
    /// <summary>
    /// Sets the dock position: top, left, right, bottom, floating
    /// </summary>
    /// <param name="pos"></param>
    virtual void setPosition(DockPosition pos) = 0;

    /// <summary>
    /// Gets the window title of the task pane
    /// </summary>
    /// <returns></returns>
    virtual std::wstring getTitle() const = 0;

    /// <summary>
    /// Registers a callback handler for task pane events such as size change.
    /// The callback will occur in Excel's main thread.
    /// </summary>
    /// <param name="events"></param>
    virtual void addEventHandler(const std::shared_ptr<ICustomTaskPaneHandler>& events) = 0;

    virtual void destroy() const = 0;
  };

  class IComAddin
  {
  public:
    using RibbonCallback = std::function<
      void(const RibbonControl&, VARIANT*, int, VARIANT**)>;
    using RibbonMap = std::function<RibbonCallback(const wchar_t*)>;

    virtual ~IComAddin() {}
    /// <summary>
    /// COM ProgID used to register this add-in
    /// </summary>
    virtual const wchar_t* progid() const = 0;
    /// <summary>
    /// Connects the add-in to Excel and processes the Ribbon XML
    /// </summary>
    virtual void connect() = 0;
    /// <summary> 
    /// Disconnects the add-in from Excel and closes any associated Ribbons
    /// </summary>
    virtual void disconnect() = 0;
    /// <summary>
    /// Sets the Ribbon XML which is passed to Excel when the <see cref="connect"/>
    /// method is called. The Office Ribbon invokes callbacks on its controls by
    /// calling the method name in the XML.  The `mapper` maps these names to 
    /// actual C++ functions.
    /// </summary>
    /// <param name="xml"></param>
    /// <param name="mapper">A map from names to functions which returns a <see cref="RibbonCallback"/> .
    /// </param>
    virtual void setRibbon(
      const wchar_t* xml,
      const RibbonMap& mapper) = 0;
    /// <summary>
    /// Invalidates the specified control: this clears the caches of the
    /// responses to all callbacks associated with the control. For example,
    /// this can be used to hide a control by forcing its getVisible callback
    /// to be invoked.
    /// 
    /// If no control ID is specified, all controls are invalidated.
    /// </summary>
    /// <param name="controlId"></param>
    virtual void ribbonInvalidate(const wchar_t* controlId = 0) const = 0;
    /// <summary>
    /// Activates the specified Ribbon pane - this method does not work for the 
    /// built-in pane
    /// </summary>
    /// <returns>true if successful</returns>
    virtual bool ribbonActivate(const wchar_t* controlId) const = 0;

    virtual std::shared_ptr<ICustomTaskPane> createTaskPane(
      const wchar_t* name,
      const IDispatch* window=nullptr,
      const wchar_t* progId=nullptr) = 0;

    using TaskPaneMap = std::multimap<std::wstring, std::shared_ptr<ICustomTaskPane>>;
    virtual const TaskPaneMap& panes() const = 0;
  };

  /// <summary>
  /// Registers a COM add-in with the specified name (which must be unique)
  /// and an option description to show in Excel's add-in options.  The add-in
  /// is not activated until you call the <see cref="IComAddin::connect"/> method.
  /// </summary>
  /// <returns>A pointer to an a <see cref="IComAddin"/> object which controls
  ///   the add-in.</returns>
  XLOIL_EXPORT std::shared_ptr<IComAddin>
    makeComAddin(const wchar_t* name, const wchar_t* description = nullptr);

  inline auto makeAddinWithRibbon(
    const wchar_t* name,
    const wchar_t* xml,
    const IComAddin::RibbonMap& mapper)
  {
    auto addin = makeComAddin(name, nullptr);
    addin->setRibbon(xml, mapper);
    addin->connect();
    return addin;
  }

  /// <summary>
  /// Converts a variant to an ExcelObj.
  /// </summary>
  /// <param name="variant"></param>
  /// <param name="allowRange">If true, a variant Range will be converted to an xlRef,
  /// otherwise it will become an array of ExcelObj.
  /// </param>
  /// <returns></returns>
  XLOIL_EXPORT ExcelObj variantToExcelObj(const VARIANT& variant, bool allowRange = false);

  XLOIL_EXPORT void excelObjToVariant(VARIANT* v, const ExcelObj& obj);
}
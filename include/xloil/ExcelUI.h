#pragma once
#include <xloil/ExportMacro.h>
#include <xloil/ExcelObj.h>
#include <xlOil/ExcelApp.h>
#include <xlOil/AppObjects.h>
#include <functional>
#include <memory>
#include <map>

typedef struct tagVARIANT VARIANT;
struct IDispatch;

namespace Excel { struct Window; }

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
    StatusBar(size_t finalTimeout = 0) : _timeout(finalTimeout) {}
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

    virtual IDispatch* content() const = 0;

    /// <summary>
    /// Gives the Window object to which this task pane is attached
    /// </summary>
    /// <returns></returns>
    virtual ExcelWindow window() const = 0;

    virtual size_t parentWindowHandle() const = 0;

    virtual void setVisible(bool) = 0;
    virtual bool getVisible() = 0;

    virtual std::pair<int, int> getSize() = 0;
    virtual void setSize(int width, int height) = 0;

    virtual DockPosition getPosition() const = 0;
    virtual void setPosition(DockPosition pos) = 0;

    virtual std::wstring getTitle() const = 0;

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
  XLOIL_EXPORT IComAddin*
    makeComAddin(const wchar_t* name, const wchar_t* description = nullptr);

  inline auto makeAddinWithRibbon(
    const wchar_t* name,
    const wchar_t* xml,
    const IComAddin::RibbonMap& mapper)
  {
    std::unique_ptr<IComAddin> addin(makeComAddin(name, nullptr));
    addin->setRibbon(xml, mapper);
    addin->connect();
    return addin.release();
  }

  XLOIL_EXPORT ExcelObj variantToExcelObj(const VARIANT& variant, bool allowRange = false);
  XLOIL_EXPORT void excelObjToVariant(VARIANT* v, const ExcelObj& obj);


  class ComAddinThreadSafe : public IComAddin
  {
  private:
    std::shared_ptr<IComAddin> _base;
  public:
    ComAddinThreadSafe(const wchar_t* name)
    {
      runExcelThread([this, nameStr = std::wstring(name)]() mutable
      {
        this->_base.reset(makeComAddin(nameStr.c_str(), nullptr));
      });
    }

    ComAddinThreadSafe(
      std::wstring&& name,
      std::wstring&& xml,
      IComAddin::RibbonMap&& mapper)
    {
      runExcelThread([this, nameStr = name, xmlStr = xml, maps = mapper]() mutable
      {
        this->_base.reset(makeComAddin(nameStr.c_str(), nullptr));
        _base->setRibbon(xmlStr.c_str(), maps);
        _base->connect();
      });
    }

    const wchar_t* progid() const override
    {
      return _base->progid();
    }
    void connect() override
    {
      connectAsync().wait();
    }
    void disconnect() override
    {
      disconnectAsync().wait();
    }
    void setRibbon(const wchar_t* xml, const RibbonMap& mapper) override
    {
      setRibbonAsync(xml, mapper).wait();
    }
    void ribbonInvalidate(const wchar_t* controlId = 0) const override
    {
      ribbonInvalidateAsync(controlId).wait();
    }
    bool ribbonActivate(const wchar_t* controlId) const override
    {
      return ribbonActivateAsync(controlId).get();
    }
    std::shared_ptr<ICustomTaskPane> createTaskPane(
      const wchar_t* name, const IDispatch* window = nullptr, const wchar_t* progId = nullptr) override
    {
      return createTaskPaneAsync(name, window, progId).get();
    }
    const TaskPaneMap& panes() const override
    {
      return _base->panes();
    }

    std::future<void> connectAsync()
    {
      return runExcelThread([obj=_base]() { obj->connect(); });
    }
    std::future<void> disconnectAsync()
    {
      return runExcelThread([obj = _base]() { obj->disconnect(); });
    }
    std::future<void> setRibbonAsync(const wchar_t* xml, const RibbonMap& mapper)
    {
      return runExcelThread([obj = _base, xmlStr = std::wstring(xml), maps = RibbonMap(mapper)]() {
        obj->setRibbon(xmlStr.c_str(), maps);
      });
    }
    std::future<void> ribbonInvalidateAsync(const wchar_t* controlId = 0) const
    {
      return runExcelThread([obj = _base,  id = std::wstring(controlId)]() { 
        obj->ribbonInvalidate(id.empty() ? nullptr : id.c_str()); 
      });
    }
    std::future<bool> ribbonActivateAsync(const wchar_t* controlId) const
    {
      return runExcelThread([obj = _base, id = std::wstring(controlId)]() { 
        return obj->ribbonActivate(id.c_str()); 
      });
    }
    std::future<std::shared_ptr<ICustomTaskPane>> createTaskPaneAsync(
      const wchar_t* name, const IDispatch* window = nullptr, const wchar_t* progId = nullptr)
    {
      return runExcelThread([obj = _base, nameStr = std::wstring(name), window]() {
        return obj->createTaskPane(nameStr.c_str(), window);
      });
    }
  };
}
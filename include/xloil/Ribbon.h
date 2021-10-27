#pragma once
#include <xloil/ExportMacro.h>
#include <xloil/ExcelObj.h>
#include <functional>
#include <memory>

typedef struct tagVARIANT VARIANT;
struct IDispatch;

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

  class ICustomTaskPaneEvents
  {
  public:
    virtual void resize(int width, int height) = 0;
    virtual void visible(bool c) = 0;
    virtual void docked() = 0;
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

    virtual intptr_t documentWindow() const = 0;
    virtual intptr_t parentWindow() const = 0;

    virtual void setVisible(bool) = 0;
    virtual bool getVisible() = 0;

    virtual std::pair<int, int> getSize() = 0;
    virtual void setSize(int width, int height) = 0;

    virtual DockPosition getPosition() const = 0;
    virtual void setPosition(DockPosition pos) = 0;

    virtual void addEventHandler(const std::shared_ptr<ICustomTaskPaneEvents>& events) = 0;
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

    virtual ICustomTaskPane* createTaskPane(
      const wchar_t* name,
      const wchar_t* progId=nullptr) const = 0;
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

  inline std::shared_ptr<IComAddin> makeAddinWithRibbon(
    const wchar_t* name,
    const wchar_t* xml,
    const IComAddin::RibbonMap& mapper)
  {
    auto addin = makeComAddin(name, nullptr);
    addin->setRibbon(xml, mapper);
    addin->connect();
    return addin;
  }

  XLOIL_EXPORT ExcelObj variantToExcelObj(const VARIANT& variant, bool allowRange = false);
  XLOIL_EXPORT void excelObjToVariant(VARIANT* v, const ExcelObj& obj);
}
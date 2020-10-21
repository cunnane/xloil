#pragma once
#include "ExportMacro.h"
#include <functional>
#include <memory>
#include <map>

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

  class IComAddin
  {
  public:
    using RibbonCallback = std::function<void(const RibbonControl&)>;
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
    /// <param name="handlers">A map from names to functions which take a 
    ///   <see cref="RibbonControl"/> argument.
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

  std::shared_ptr<IComAddin> xloil::makeAddinWithRibbon(
    const wchar_t* name,
    const wchar_t* xml,
    const IComAddin::RibbonMap& mapper)
  {
    auto addin = makeComAddin(name, nullptr);
    addin->setRibbon(xml, mapper);
    addin->connect();
  }
}
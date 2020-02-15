# xlOil

xlOil is a framework for building Excel language bindings. That is, a way to write functions in a language of your choice and have them appear in Excel.

xlOil is designed to have very low overheads when calling your own worksheet functions.

xlOil supports different languages via plugins. The languages currently supported are:

- C++
- Python

You can use xlOil as an end-user of these plugins or you can use it to write you own language bindings and contribute.

## Getting started (user)

- Requirements:
  - Excel 2010 (64-bit) or newer
  - For xlOil_Python: Python 3.6 or 3.7 with numpy
- Extract files from the latest release zip to a single directory
  - You need the `xloil.xll`, `xlOil_Core.dll` and `xloil.ini` files
  - For the Python plugin, you need `xlOil_PythonXY.dll`, `xloil.py` and `xlOil_PythonXY.ini` where XY is the Python version you want to use
- (Optional) Edit the ini files
  - Choose which plugins to load in `xloil.ini`
  - For the Python plugin, you may need to set the correct Python paths
- Drop `xloil.xll` into an Excel session
- Read the docs for the plugins you want to use
- Enjoy!
- If the addin fails to load or something else goes wrong, look for errors in `xloil.log` in the same directory as `xloil.xll`

To install the add-in so it starts with Excel, place all the files in your XLSTART folder.  You can rename `xloil.xll` and `xloil.ini` if desired - just ensure they have the same base-name.

### Editing the ini file

They are actually a TOML files. The settings are explaing in comments. The main choice is whether to load specific plugins with:

    Plugins=["xloil_Python37.dll", "foobar.dll"]

Or search for plugins in the same directory as xloil.xll with:

    PluginSearchPattern="xloil_*.dll"

You can specifiy both.

Plugin ini files can contain a special section called `[Environment]`.  All name=value pairs in this block are set in the order specifed as environment variables *before* the plug-in is loaded. Environment strings in %...% are expanded.  Writing name="<HKLM\RegKey\Value>" will fetch the requested key from the registry.

## Getting started (xlOil developer)

- You need Visual Studio 2017 or newer
- All xlOil_Core dependencies are already in the `external` folder
- For xlOil_Python you need to set the correct paths in `PySettings.props`, `PySettings36.props` and `PySettings37.props`
- For debugging, set xlOil_Loader as the target project, command=`<Path-to-Excel.exe>` args=`$(OutDir)\xloil.xll`

## Why write xlOil

This is for people thinking about writing language bindings. If you want to write worksheet functions in a nice language, skip to the plugin documentation.

Interfacing with Excel is tricky for a general language. You have a choice of poisons:

- C-API - is C and hence unsafe, the API is also old, has some quirks and is missing many features
- COM - more fully-featured but slower and missing some features of C-API. Requires COM binding support in your language or a great deal of pain will be endured
- .Net API - actually sits on top of COM, good but limited to .Net languages

xlOil tries to give you the first two blended in a more friendly fashion and adds:

- Solution to the "how to register a worksheet function without a static DLL entry point" problem
- Object caching
- A framework for converting excel variant types to another language and back
- A loader stub
- Goodwill to all men

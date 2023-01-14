=========================
xlOil C++ Getting Started
=========================

There are a few ways to build a C++ project with xlOil:

   1. xlOil plugin: writing a DLL which can be loaded by xlOil after the loader XLL 
      starts up the application. This allows co-operation with other xlOil features such as
      the the object cache and settings file.  It means that cache references created by 
      `xloRef` can be used in your addin.
   2. XLL dynamically linking xlOil. This allows co-operation with the Excel object cache.
      Your XLL will need to be able to find xlOil when it loads - this may be easier if delay
      loading is employed.
   3. XLL statically linking xlOil. This gives a standalone XLL. It does not include
      the xlOil cache functions `xloRef` and `xloVal`, although it is easy to define a private
      object cache with your own function names.

The header files and `xlOil.lib` file required for dynamic linking are included in the 
`binary releases <https://gitlab.com/stevecu/xloil/-/releases>`_.

You need: 
    * `xlOil.lib`: linking library
    * `xlOil.dll`: main binary 
    * `xlOil.xll`: XLL loader stub
    * `xlOil.ini`: Main settings file
    * `include/xloil`: header files

The static libs are a lot larger, so you'll need to build them from source.

If you want to read settings from the add-in's *.ini* file using *tomlplusplus* you can find 
this in `xlOil/external` - it's header-only.

Creating an xlOil Plugin
~~~~~~~~~~~~~~~~~~~~~~~~

This method assumes you already have a running install of xlOil.

Create a project & solution for your DLL.

Building an xlOil plugin requires a number of build settings and the easiest way to get these right
it is use a property sheet.  First, include an edited copy of `./libs/SetRoot.props` where *xlOilRoot* 
points to the xlOil directory. Then include `./libs/xlOilPlugin.props`.

Now add the following source file:

.. highlight:: c++

:: 

    #include <xloil/xloil.h>

    XLO_PLUGIN_INIT(xloil::AddinContext* ctx, const xloil::PluginContext& plugin)
    {
      xloil::linkLogger(ctx, plugin);
      return 0;
    }

    using namespace xloil;

    XLO_FUNC_START( 
        MyFunc(const ExcelObj* arg1, const ExcelObj* arg2)
    )
    {
        auto result = arg1->toString() + ": " + arg2->toString(L"default value");
        return ExcelObj::returnValue(result);
    }
    XLO_FUNC_END(MyFunc).threadsafe()
        .help(L"Joins two strings")
        .arg(L"Val1", L"First String")
        .arg(L"Val2", L"Second String");


Declaring an `XLO_PLUGIN_INIT` entry point lets xlOil know the DLL is a plugin. Calling `linkLogger`
links the *spdlog* to the main xlOil log output.  This is optional if you don't plan on using any logging
functionality.

The ``ExcelObj`` class wraps the xloper12 variant type used in the XLL interface. It provides
a number of accessors with a focus on efficiency and avoidance of copying (see below).

To load your plugin in Excel you need to ensure that the main xlOil add-in can locate your DLL. 
You can do this by changing the Addin enviroment in the `xlOil.ini` file or changing the
current working directory or system PATH, or by simply ensuring your DLL is in the same directory
as *xlOil.dll*.

You should also add your plugin to the `Plugins=` key in `xlOil.ini`.


Creating an XLL which statically links xlOil
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Build xlOil with the *DebugStatic* or *ReleaseStatic* configuration.

Create a project & solution for your XLL. It should output a DLL, but change the *Target Extension* to XLL.

Building an xlOil-linked XLL requires a number of build settings and the easiest way to get these right
it is use a property sheet.  First, include an edited copy of `./libs/SetRoot.props` where *xlOilRoot* 
points to the xlOil directory. Then include `./libs/xlOilStaticLib.props`.

Alternatively, you can copy and modify the `TestAddin` project, though you will still need to ensure
*xlOilRoot* is set correctly.

Create a `Main.cpp` file with the following contents:

.. highlight:: c++

::

    #include <xloil/xlOil.h>
    #include <xloil/XllEntryPoint.h>
    using namespace xloil;
    
    struct MyAddin
    {};
    XLO_DECLARE_ADDIN(MyAddin);

    XLO_FUNC_START( 
        MyFunc(const ExcelObj* arg1, const ExcelObj* arg2)
    )
    {
        auto result = arg1->toString() + ": " + arg2->toString(L"default value");
        return ExcelObj::returnValue(result);
    }
    XLO_FUNC_END(MyFunc).threadsafe()
        .help(L"Joins two strings")
        .arg(L"Val1", L"First String")
        .arg(L"Val2", L"Second String");


Debugging your addin
~~~~~~~~~~~~~~~~~~~~

Start a debugging session with the following settings where the XLL may be `xlOil.xll`
or one you built yourself depending on which path you are following.

   * Command: <path to Excel.exe>
   * Arguments: <path to the XLL>

Setting the working directory is optional, but it may help locate any externals DLLs your
add-in uses.

There are many examples to follow in the ``xloil_Utils`` and ``xloil_SQL`` projects.
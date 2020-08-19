======================
xlOil C++
======================

The xlOil C++ interface has a `doxygen API description <doxygen/index.html>`_. It provides
a layer above the `XLL API <https://docs.microsoft.com/en-us/office/client-developer/excel/excel-xll-sdk-api-function-reference>`_
and includes features not available via the XLL API, such as RTD Servers and event handling. 


Getting Started
----------------

Create a C++ DLL project with the following source file:

.. highlight:: c++

:: 

    #incude <xloil/xloil.h>

    XLO_PLUGIN_INIT(xloil::AddinContext* ctx, const xloil::PluginContext& plugin)
    {
      xloil::linkLogger(ctx, plugin);
      return 0;
    }

This lets xlOil know the DLL is a plugin and links the *spdlog* instance local to 
the addin DLL to the xlOil log output.

To create an Excel function add the following to a source file:

.. highlight:: c++

:: 

    #incude <xloil/xloil.h>
    using namespace xloil;

    XLO_FUNC_START( 
        MyFunc(const ExcelObj* arg1, const ExcelObj* arg2)
    )
    {
        auto result = arg1->toString() + ": " + arg2->toString();
        return ExcelObj::returnValue(result);
    }
    XLO_FUNC_END(MyFunc).threadsafe()
        .help(L"Joins two strings")
        .arg(L"Val1", L"First String")
        .arg(L"Val2", L"Second String");

The ``ExcelObj`` class wraps the xloper12 variant type used in the XLL interface. It provides
a number of accessors with a focus on efficiency and not copying data.

There are many examples to follow in the ``xloil_Utils`` and ``xloil_SQL`` projects.

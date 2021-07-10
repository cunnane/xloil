=======================
xlOil C++ Using COM
=======================

The root of Excel's COM interface is `Excel.Application <https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)>`_
which VBA programmers will be familiar with.

The easiest way to work with it in C++ is by importing the type library into your source file - 
this allows compile time checks and auto-completion and results in faster code.  COM objects
can be accessed using 'late-binding' where methods and properties are queried from the object
at runtime, but this clearly limits performance.

In xlOil, the appropriate type library can be imported with the *ExcelTypeLib.h* header:

::

    #include <xlOil\ExcelTypeLib.h>
    #include <xlOil\xlOil.h>
    using namespace xloil;

    XLO_FUNC_START(testToday())
    {
        excelApp().Range[L"A1"]->NumberFormat = L"dd-mm-yyyy";
    }
    XLO_FUNC_END(testToday).macro();

If *ExcelTypeLib.h* is included before `xlOil.h` it adds COM error handling to any worksheet 
functions declared with `XLO_FUNC_START`.  This is useful because COM throws its own `_com_error`
classes which do not derive from `std::exception`.

I cannot find a specific Excel/COM/C++ tutorial at this time, but the *Excel.Application* API 
documentation is comprehensive. 

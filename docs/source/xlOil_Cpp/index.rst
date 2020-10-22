======================
xlOil C++
======================

xlOil C++ provides a layer above the `XLL API <https://docs.microsoft.com/en-us/office/client-developer/excel/excel-xll-sdk-api-function-reference>`_
and includes features not available via the XLL API, such as RTD Servers, Ribbon customisation
and event handling.

The xlOil C++ interface has a `doxygen API description <doxygen/index.html>`_.

.. toctree::
    :maxdepth: 4
    :caption: Contents

    GettingStarted
    SpecialArgs
    Events
    CustomGUI


Quick Tour of the Key Classes
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

ExcelObj - An XLOPER wrapper
----------------------------

An *ExcelObj* is a friendly wrapper around Excel's *xloper12* struct which is the variant type
used to pass arguments in the XLL interface.  An *ExcelObj* it has the same size and *xloper12* so
you can freely cast between the two (although be aware that *ExcelObj* has a destructor).
*ExcelObj* supports a number of variant-like operations: it can be compared to strings or numbers,
converted to a string, and created from values or initialiser lists. However, it's not recommended
to use *ExcelObj* as a generic variant type - reserve it for interaction with the XLL API.

ExcelArray / ArrayBuilder
-------------------------

To **view** a suitable `ExcelObj` as an array, pass it to the `ExcelArray` constructor.  
By default, the array is 'trimmed' to last row and column containing data (i.e. not empty or 
N/A).  `ExcelArray` is a view - it is lightweight but does not own the underlying
data.

The `ArrayBuilder` class constructs arrays to return to Excel. For efficiency (of the CPU not
the programmer!) it requires you to know up-front how many array elements you require and the
total length of all strings in the array.  This may mean you need to make two passes of your 
data, but it saves iterating through the array on destruction. 

PString
-------

Strings in the XLL API are Pascal wide char strings: the first character is the length and 
they are not null terminated.  `PString` wraps such a string and provides a similar interface to 
`std::wstring`.  There's also a `PStringView` type like `std::wstring_view`.

Range / RangeArg / ExcelRef
---------------------------

There are few types of range class with slightly different internals and usages. All behave 
in a "range-like" way, supporting:

    * `(i, j)` - accessor to pick out individual elements
    * `range(fromRow, fromCol, toRow, toCol)` - create a sub-range
    * `value()` - convert to an `ExcelObj` single value or array
    * `address()` - fetch the sheet address as a string 
    * `set(val)` - sets the range values to the given `ExcelObj` (which may be an array)


The `RangeArg` type should only be used to declare an argument to a user-defined worksheet
function.  Specifing this type means that xlOil is allowed to pass range references to your
function.  See :any:`SpecialArgs`.  A `RangeArg` is not directly constructable - you should
convert it to an `ExcelRef` to copy and pass it around.

The `ExcelRef` type is equivalent to an *xltypeRef* `ExcelObj`, that is it points to a 
rectangular space on a specific sheet.  It can be created from an `ExcelObj` which is a 
range reference (typically via `RangeArg.toExcelRef()`) or a sheet address string.

The `Range` type is a virtual base: it may point either to an `ExcelRef` range or a COM 
range object.  Prefer `ExcelRef` when you are sure the range information has been passed via
an XLL worksheet function.


excelCall - calling XLL API functions
-------------------------------------

The XLL SDK docs describe a number of API calls beginning with `xl` such as `xlCoerce` and
`xlSheetNm`. There are also many more poorly documented API calls beginning with `xlc` and 
`xlf`. They can all be found in `xlcall.h`.  Invoking these functions is straightforward:

.. highlight:: c++

::
    
    ExcelObj result = excelCall(msxll::xlSheetId, "Sheet1");


All arguments are automatically converted to `ExcelObj` type where a conversion is possible
before being passed to Excel.  The type of the return value is given in the documentation.


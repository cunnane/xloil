==============
xlOil Concepts
==============

This document explains some of the key xlOil concepts shared accross different languages

.. contents::
    :local:


Excel Functions (UDFs)
----------------------

Excel supports several classes of user-defined functions:

- Macros: run at user request, have write access to workbook
- Worksheet functions: run by Excel's calculation cycle. Several sub-types:
  - Vanilla
  - Thread-safe: can be run concurrently
  - Macro-type: can read from sheet addresses and invoke a wider variety of Excel interface functions
  - Async: can run asynchronously during the calc cycle, but not in the background
  - RTD: (real time data) background threads which push data onto the sheet when it becomes available
  - Cluster: can be packaged to run on a remote Excel compute cluster

xlOil currently supports all but Cluster functions.

Excel can pass functions / macros data in one of these types:

- Integer
- Boolean
- Floating point
- String
- Error, e.g. #NUM!, #VALUE!
- Empty
- Array of any of the above
- Range refering to a worksheet address

There is no date type. Excel's builtin date functions interpret numbers as days since 1900. 
Excel does not support timezones.


.. _core-cached-objects:

Cached Excel Objects
--------------------

xlOil has an internal lookup for Excel values, which is a convenient way of 
passing arrays around a sheet and as arguments to other xlOil functions.

The function ``=xloRef(A1:B2)`` returns a cache string of the form:
``<UniqueChar>[WorkbookName]SheetName!CellRef,#``

The data can be recovered using ``=xloVal(<CacheString>)``. Alternatively,
this string can be passed instead of the source range to xlOil functions which
support cache lookups - and for large arrays this is much faster.

Example use cases are:

    * Where you might use a named range - to avoid updating references 
      in multiple functions when data is appended.  Because `xloRef` automatically 
      trims the range back to the last non-blank row, it can be pointed to a range
      far larger than the data.
    * To speed up passing the same large array into several functions 
      (e.g. multiple lookups from a data table, although consider xlOil_SQL if
      you want to do this).

However, there is a disadvantage to using `xloRef`: the cache is cleared when
a workbook is closed, but Excel does not know to recalculate the `xloRef` 
functions when the workbook is reopened. Hence you need to force a sheet
recalculation using *Ctrl-Alt-F9*.

In addition to caching arrays, xlOil plugins use the cache to opaquely return
referencs to in-memory structures.  Although the strings look similar, they 
point to very different objects and cannot be written to the sheet using `xloVal`.


.. _concepts-rtd-async:

Rtd / Async
-----------

In Excel, RTD (real time data) functions are able to return values independently of Excel's 
calculation cycle.  Excel has supported RTD functions since at least Excel 2002.  In Excel 
2010, Excel introduced native async functions.

RTD:

    * Pro: operates independently of the calc cycle - true background execution
    * Pro: provides notification when an RTD function call is changed or removed
    * Con: increased overhead compared to native async
    * Con: requires automatic calculation enabled (or repeated presses of F9 until calc is done)

Native async:

    * Pro: Less overhead compared to RTD
    * Pro: works with manual calc mode
    * Con: tied to calc cycle, so any interruption cancels all asyncs functions

The last con is particularly problematic for native async: *any* user interaction with Excel will
interrupt the calc, so whilst native async functions can run asynchronously with each other, they
cannot be used to perform background calculations.

.. _concepts-ribbon:

Custom UI: The Fluent Ribbon
----------------------------

xlOil allows dynamic creation of Excel Ribbon components. The ribbon is defined by XML
(surrounded with <customUI> tags) which should be created with a specialised editor, see the 
*Resources* below. Controls in the ribbon interact with user code via callback handlers.  
These callbacks pass a variety of arguments and may expect a return value; it is important 
to check that the any callback behaves as per the callback specifications in the *Resources*.

To pass ribbon XML into Excel, xlOil creates a COM-based add-in in addition to the XLL-based 
add-in which loads the xlOil core - you can see this appearing in Excel's add-in list in the 
Excel's Options windows.

Resources:

   * `Microsoft: Overview of the Office Fluent Ribbon <https://docs.microsoft.com/en-us/office/vba/library-reference/concepts/overview-of-the-office-fluent-ribbon>`_
   * `Microsoft: Customizing the Office Fluent Ribbon for Developers <https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/aa338202(v=office.12)>`_
   * `Microsoft: Custom UI XML Markup Specification <https://docs.microsoft.com/en-us/openspecs/office_standards/ms-customui/31f152d6-2a5d-4b50-a867-9dbc6d01aa43>`_
   * `Microsoft: Ribbon Callback Specifications <https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ee691833(v=office.14)>`_
   * `Office RibbonX Editor <https://github.com/fernandreu/office-ribbonx-editor>`_
   * `Ron de Bruin: Ribbon Examples files and Tips <https://www.rondebruin.nl/win/s2/win003.htm>`_
   

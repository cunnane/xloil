=========================
xlOil Python Concepts
=========================

.. contents::
    :local:

Cached Objects
--------------

If xlOil cannot convert a returned python object to Excel, it will place it in 
a object dictionary and return a reference string of the form
``<UniqueChar>[WorkbookName]SheetName!CellRef,#``

If a string of this kind if encountered when reading function arguments, xlOil 
automatically tries to fetch the corresponding python object. With this
mechanism you can pass python objects opaquely between functions by using
cache reference strings outputted from one function as arguments to another.

For example:

::

    @xlo.func
    def make_lambda(pow):
        return lambda x: x ** pow

    @xlo.func
    def apply_lambda(f, x):
        return f(x)

Since xlOil doesn't know how to convert a lambda function to an Excel object,
the first function will output a cache reference string.  That string will be 
automatically turned back into a lambda if passed as an argument to the second 
function.

The python cache is separate to the Core object cache accessed using `xloRef`
and `xloVal`.  The Core cache stores native Excel objects such as arrays.
When reading functions arguments xlOil tries to lookup strings in both of these
caches. 

The leading `UniqueChar` means xlOil can very quickly determine that a string
isn't a cache reference, so the overhead of checking if every string argument
is a cache object is very fast in practice. 


Local Functions and Workbook Modules
------------------------------------

When an Excel workbook is opened, xlOil tries to load the module `<workbook_name>.py` 
(this is configurable).

When registering functions from such a workbook module, xlOil defaults to making
any declared functions "local". This means their scope is limited to the workbook.
(It achieves this by creating a VBA stub to invoke them). This scoping can be
overidden by the `xlo.func` decorator.

Local functions have some limitations compared to global scope ones:
- No native async or threadsafe, but RTD async is OK
- Slower due to the VBA redirect
- Workbook must be saved as macro-enabled (xlsm extension)
- No function wizard help, but CTRL+SHIFT+A to show argument names is available

Another way to package python code for distribution is to create an XLL, see
:ref:`core-distributing-addins`

xlOil sets the module-level variable `_xl_this_workbook` to the workbook name in a 
workbook module.

(Technical note: It is possible to use the Application.MacroOptions call to add help to the 
function wizard for VBA, but identically named functions will conflict which somewhat defeats 
the purpose of local functions).


Array Functions
---------------

By default, xlOil-Python converts Excel array arguments to numpy arrays. The conversion
happens entirely in C++ and is very fast.  Where possible you should write functions
which support array processing (vectorising) to take advantage of this, for example
the following works when it's arguments are arrays or numbers:

::

    @xlo.func
    def cubic(x, a, b, c, d)
        return a * x ** 3 + b * x ** 2 + c * x  + d

You can take this further to evaluate polynominals of n-th degree:

::

    @xlo.func
    def poly(x, coeffs: xlo.Array(float)):
        return np.sum(coeffs * X[:,None] ** range(coeffs.T.shape[0]), axis=1)

Specifing the type of the array avoids xlOil needing to scan the element to determine it.
There is a problem with this function: what happens if ``x`` is two-dimensional?  To avoid
this quandry we can specify:

::

    @xlo.func
    def poly(x: xlo.Array(dims=1), coeffs: xlo.Array(float)):
        return np.sum(coeffs * X[:,None] ** range(coeffs.T.shape[0]), axis=1)



Dates
-----

In the Python plugin, just applying the argument annotation `datetime` will request a date 
conversion. Dates returned from functions will be converted to Excel date numbers:

::

    from datetime import datetime, timedelta
    @func
    def AddDay(date: datetime):
        return date + timedelta(days = 1)


xlOil can interpret strings as dates. In the settings file, the key ``DateFormats`` 
specifies an array of date-formats to try when parsing strings. Naturally, adding more 
formats decreases performance.  The formats use the C++ ``std::get_time`` syntax,
see https://en.cppreference.com/w/cpp/io/manip/get_time.

Since ``std::get_time`` is **case-sensitive** on Windows, so is xlOil's date parsing
(this may be fixed in a future release as it is quite annoying for month names).

Excel has limited internal support for dates. There is no primitive date object in the  
XLL interface used by xlOil, but cells containing numbers can be formatted as dates.  
This means that  worksheet functions cannot tell whether values received are intended 
as dates - this applies to Excel built-in date functions as well: they will interpret 
any number as a date. (It is possible to check for date formatting via the COM interface 
but this would give behaviour inconsistent with the built-ins)

Excel does not understand timezones and neither does ``std::get_time``, so these
are currently unsupported.


Events
------

With events, you can request a callback on various user interactions. If you are familiar  
with VBA, you may have used Excel's event model already.  Most of the workbook events 
described in `Excel.Appliction <https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)#events>`_
are available in xlOil. See the xloil.Event documention for the complete list.

Some events use reference parameters, for example setting the `cancel` bool in `WorkbookBeforeSave`, 
cancels the event.  In xlOil you need to set this value using `cancel.value=True` as python 
does not support reference parameters for primitive types.

Events are (currently) global to the Excel instance, so you may need to filter by workbook name when 
handling events.

xlOil has some extra events:

    * `WorkbookAfterClose`: Excel's event *WorkbookBeforeClose*, is cancellable by the user so it is 
      not possible to know if the workbook actually closed. `WorkbookAfterClose` fixes this but there
      may be a long delay before the event is fired.
    * `CalcCancelled`: called when the user interrupts calculation, maybe useful for async functions

Examples
~~~~~~~~

::

    def greet(workbook, worksheet):
        xlo.Range(f"[{workbook}]{worksheet}!A1") = "Hello!"

    xlo.event.WorkbookNewSheet += greet


Looking for xlOil functions in imported modules
-----------------------------------------------

To tell xlOil to look for functions in a python module use ``xloil.scan_module(name)``. 
xlOil will import ``name`` if required, then look for decorated functions to register.


xloPyLoad: import and scan a python module (worksheet function)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

.. function:: xloPyLoad(ModuleName)

    Imports the specifed python module and scans it for xloil functions by calling
    ``xloil.scan_module(name)``

============
xlOil Python
============

.. toctree::
    :maxdepth: 4
    :caption: Contents

    xlOil_Python_Example
    xlOil_Python_Module
    xlOil_Python_Jupyter
    xlOil_Python_Rtd

Introduction
------------

The Python plugin for xlOil allows creation of Excel functions and macros backed by Python
code.

xlOil_Python loads specified python modules, looking for functions to add to Excel's global
name scope, like an Excel addin.  The plugin can also look for modules of the form
<workbook_name>.py and load these too, this is like creating a VBA code module for a workbook.
Any python module which contains Excel functions is watched for file modifications so 
code changes are reflected immediately in Excel.

xlOil_Python is tightly integrated with numpy, allowing creation of fast Excel array 
functions.

For examples, have a look at :doc:`xlOil_Python_Example` and :ref:`core-example-sheets`.

Getting Started
---------------

Run the following at a command prompt with python environment settings:

::

    pip install xlOil
    xloil install

The call to ``xloil install`` registers the xlOil addin with Excel and places a settings
files at the `%APPDATA%/xlOil/xlOil.ini`.  The settings file describes the python modules 
which will be loaded and sets the paths to the python libraries binaries. xlOil attempts 
to set the python paths automatically using data in the Windows registry, but can be 
overriden if required.

To test the setup, you can try the python example sheet: :ref:`core-example-sheets`.

Note it's not necessary for ``xloil.xll`` to be registered as an addin: you can just drop
it into your Excel session when required. 

You now have several ways to get xlOil to load your python:


My first xlOil module
~~~~~~~~~~~~~~~~~~~~~

Now let's make our first python module using xlOil.  Create a `MyTest.py` file with 
the following lines:

::

    import xloil as xlo

    @xlo.func
    def Greeting(who):
        return "Hello  " + who

Edit `%APPDATA%/xlOil/xlOil.ini` so that `PYTHONPATH` includes the 
directory containing `MyTest.py` then add `MyTest` to the `LoadModules` key.

Now open Excel and call the Greeting function.

My first xlOil addin
~~~~~~~~~~~~~~~~~~~~~

We might like to distribute our code as a packaged addin so users don't have 
to edit `xlOil.ini`. To do this, run the following at a command prompt:

::

    xloil create myaddin.xll

This will create a `myaddin.xll` and `myaddin.ini` in the current directory.
By default, the XLL will try to load `myaddin.py`, so let's create it:

::

    import xloil as xlo

    def MySum(x, y, z):
        '''Adds up numbers'''
        return x + y + z

Now drop `myaddin.xll` into an Excel session and try to use ``MySum``.

For more on this packaging addins, see :ref:`core-distributing-addins`.

My first workbook module
~~~~~~~~~~~~~~~~~~~~~~~~

Create and an Excel workbook called `MyBook`. In the same directory, create 
a file `MyBook.py` containing the following:

::

    import xloil as xlo

    @xlo.func
    def Adder(x, y):
        return x + y

You'll need to open and close `MyBook` in Excel for xlOil to find the python file.
Now try invoking the Adder function - it can also add arrays!

If this isn't working, ensure that "Trust accesst to the VBA object model" 
is checked in Excel Options -> Trust Centre -> Macro Settings.


Getting Started (trouble)
-------------------------

Check the `xlOil.log` file for errors. By default, the log file is created in the
same directory as `xlOil.ini` in your AppData directory.  If xlOil core has 
succesfully loaded, the worksheet function `xloLog` will tell you where this file is.

You may need to set the python paths in the `xlOil.ini` file for xlOil to find 
your python distribution.


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

Excel has limited support for dates. There is no primitive date object in the XLL
interfae, but cells containing numbers can be formatted as dates.  This means that 
worksheet functions cannot tell whether values received are intended as dates - this 
applies to Excel built-ins as well: they will interpret any number as a date if
requested.  (It is possible to check for dates via the COM interface but this would
give behaviour inconsistent with the built-ins)

In the Python plugin, just applying the argument annotation `datetime.datetime`
will request a date conversion.

xlOil can interpret strings as dates. In the settings file, the key 
``DateFormats`` specifies an array of date-formats to try when parsing strings.
Naturally, adding more formats decreases performance.  The formats use the
``std::get_time`` syntax, see https://en.cppreference.com/w/cpp/io/manip/get_time.
Since ``std::get_time`` is **case-sensitive** on Windows, so is xlOil's date parsing
(this may be fixed in a future release).

Excel does not understand timezones and neither does ``std::get_time``, so these
are currently unsupported.


Events
------

With events, you can request a callback on various user interactions. If you are familiar  
with VBA, you may well have come across Excel's event model before.  Most of the workbook events 
described in `Excel.Appliction <https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)#events>`_
are availble in xlOil. See the xloil.Event documention for the list.

Some events use reference parameters, for example setting the `cancel` bool in `WorkbookBeforeSave`, 
cancels the event.  In xlOil Iyou need to set the value using `cancel.value=True` because python 
does not support reference parameters for primitive types.

Events are (currently) global in the Excel instance, so you may need to check the workbook name when 
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

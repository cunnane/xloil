=========================
xlOil Python Introduction
=========================

.. contents::
    :local:

Introduction
------------

The Python plugin for xlOil primarily allows creation of Excel functions and macros 
backed by Python code. In addition it offers full control over GUI objects and an 
interface for Excel automation: driving the application in code.

When loaded by xlOil, *xlOil_Python* loads specified python modules, looking for functions 
to add to Excel's global name scope, like an Excel addin.  The plugin can also look for modules 
of the form <workbook_name>.py and load these too, this is like creating a VBA code module for 
a workbook. Any python module which contains Excel functions is watched for file modifications so 
code changes are reflected immediately in Excel.

*xlOil_Python* is tightly integrated with numpy, allowing creation of fast Excel array 
functions.

*xlOil_Python* can be imported in python code to allow remote control of an Excel application.
See :doc:`ExcelApplication`

For examples of worksheet functions and GUI controls have a look at :doc:`Example` and
:ref:`core-example-sheets`.

Getting Started
---------------

**You must use the same bit-ness of python and Excel**.  So if your Excel is 32-bit, you must
install xloil using a 32-bit python distibution.

Run the following at a command prompt with python environment settings:

::

    pip install xlOil
   
xlOil can now be imported to allow remote control of an Excel application.  See :doc:`ExcelApplication`

To install the addin which allows you to create python-based Excel functions, type:

::

     xloil install

This call registers the xlOil addin with Excel and places a settings file at
`%APPDATA%/xlOil/xlOil.ini`.  The settings file describes the python modules 
which will be loaded and sets the paths to the python distribution. xlOil should 
set the python paths automatically, but they can be overriden if required.

To test the setup, you can try the python example sheet: :ref:`core-example-sheets`.

.. note:: 
    It's not necessary for ``xlOil.xll`` to be registered in this way: you can just
    drop it into your Excel session when required. 

You now have three ways to get xlOil to load your python code.


My first xlOil module
~~~~~~~~~~~~~~~~~~~~~

Create a `MyTest.py` file with the following lines:

::

    import xloil as xlo

    @xlo.func
    def Greeting(who):
        return "Hello  " + who

Open Excel and use the *xlOil* ribbon toolbar to ensure the search paths include
the directory containing `MyTest.py`.  *Then* add 'MyTest' to the loaded modules.
(the order matters because editing the *Load Modules* triggers a load of all newly
added modules)

Call the `=Greeting("world")` function in a cell.


My first workbook module
~~~~~~~~~~~~~~~~~~~~~~~~

Create an Excel workbook called `MyBook`. In the same directory, create a
file `MyBook.py` containing the following:

::

    import xloil as xlo

    @xlo.func
    def Adder(x, y):
        return x + y

You need to open and close `MyBook` in Excel for xlOil to find the python file.
Now try invoking the `=Adder()` function - it can also add arrays!

If this isn't working, ensure that "Trust access to the VBA object model" is
checked in *Excel Options -> Trust Centre -> Macro Settings* - this setting
is off by default.

Using the Ribbon
----------------

xlOil's Ribbon toolbar can:

    * Change the python environment (reqires restart)
    * Select modules to load at startup and *sys.path* to set
    * Open the log file
    * Open a console to interact with the embedded python environment
    * Choose a debugger, see :doc:`/xlOil_Python/Debugging`
    * Select date formats to use when parsing strings

The toolbar edits the settings file so that changes persist.  The ribbon is enabled by
but can be disabled by removing it from the specified *Load Modules*.

.. note::

    If you have an old ini file (prior to v0.15), you will need to upgrade it to use the  
    ribbon toolbar. Remove the old ini file and remove/install xlOil.

Troubleshooting
---------------

If xlOil detects a serious load error, it pops up a log window to alert you (this can
be turned off). If it succesfully loaded the core DLL a log file will also be created
in `%APPDATA%/xlOil` next to `xlOil.ini`.  The worksheet function `=xloLog()` will tell 
you where this file is.

Normally a python distribution or environment can be loaded with only the location of 
*python.exe* passed via the `PYTHONEXECUTABLE` environment varaible.  For more complex
setups, you may need to set the python paths, i.e. `PATH` and `PYTHONPATH` and maybe even 
`PYTHONHOME`, in the `xlOil.ini` file for xlOil to load your python distribution.

If the xlOil ribbon does not appear, check that `xloil.xloil_ribbon` appears in the
*LoadModules* key in the ini file.

Intellisense / Function Context Help
------------------------------------

To activate pop-up function help, follow the instructions here: :any:`concepts-intellisense`.

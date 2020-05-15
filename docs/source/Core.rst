===========
xlOil
===========

xlOil is a framework for building Excel language bindings. That is, a way to 
write functions in a language of your choice and have them appear in Excel.

xlOil is designed to have very low overheads when calling your own worksheet 
functions.

xlOil supports different languages via plugins. The languages currently 
supported are:

- C++
- Python
- SQL

You can use xlOil as an end-user of these plugins or you can use it to write
you own language bindings and contribute.

Getting Started
---------------

Currently you need **Excel 2010 (64-bit)** or later.

If you have python available, you can install xlOil via pip. See 
:doc:`xlOil_Python`, otherwise follow the below instructions.

Download the binary package (e.g. from gitlab) and unzip to a directory of 
your choice. 

You can run the `xlOil_Install.ps1` script to install the addin for every
Excel session, or just drop `xloil.xll` into any running Excel session
to load xlOil temporarily.

xlOil should now load when you open Excel, try following 
:any:`xlOil_SQL/Getting Started`

To configure the plugins being loaded, see :any:`Edit the settings files`.

Troubleshooting
~~~~~~~~~~~~~~~

If the addin fails to load or something else goes wrong, look for errors in 
`xloil.log` created in the same directory as `xloil.xll`. You may need to 
:any:`Edit the settings files`

.. _core-example-sheets:

Example sheets
--------------

To check your setup and see some of the capabilities of xlOil, try:
:download:`Tests and Examples </../build/xlOilExamples.zip>`.

Edit the settings files
-----------------------

There is an `xlOil.ini` file linked to the main xloil.xll addin. (This ini file 
is actually parsed as TOML, an extension of the ini format).

xlOil first searches for this file in the directory containing the `xlOil.xll` 
addin, then in `%APPDATA%/xlOil`.

The most important setting in `xlOil.ini` is the choice of plugins:

::

    Plugins=["xloil_Python.dll", "xlOil_SQL.dll"]

xlOil searches for these plugins first in the directory containing `xlOil.xll`, 
then the usual DLL search paths. Optionally you can load all plugins in the 
same directory as `xlOil.xll` with a pattern match, by setting:

::

    PluginSearchPattern="xloil_*.dll"

xlOil won't load the same plugin twice if these two methods overlap!

Setting enviroment variables in settings files
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Each plugin can have an *Environment* section in the settings file. Within this block
tokens are interpreted as enviroment variables to set. A plugin's environment settings 
are processed before the plugin is loaded. 

Keys are interpreted as environement variables to set. Values can reference other enviroment 
variables by surrounding the name with `%` characters.

In addition you can pull values from the registry by surrounding the registry
path with angle brackets `<>`. Leaving a trailing backslash `\\` in the 
registry path fetches the default value for that key.

The TOML syntax of three single quotes indicates a string literal, this avoids escaping 
all the backslashes.

The default enviroment block for Python looks like this:

::

    [[xlOil_Python.Environment]]
    xlOilPythonVersion="3.7"

    [[xlOil_Python.Environment]]
    PYTHONPATH='''<HKLM\SOFTWARE\Python\PythonCore\%xlOilPythonVersion%\PythonPath\>'''
    PYTHON_LIB='''<HKLM\SOFTWARE\Python\PythonCore\%xlOilPythonVersion%\InstallPath\>'''

    [[xlOil_Python.Environment]]
    PATH='''%PATH%;%PYTHON_LIB%;%PYTHON_LIB%\Library\bin'''

The double brackets tell TOML that the order of these declarations is important,
this means we can refer to previously set enviroment variables.

The log file
------------

If a function doesn't appear or behave as expected, check the log file created by default
in the same directory as xlOil.xll.  A setting in `xloil.ini` controls the log level.

A common problem is that the COM interface misbehaves either failing on start-up or failing
because of an open dialog box in Excel.  For a start-up fail, unload and reload the addin. 
For other errors try to close dialog boxes or panes and if that fails, restart Excel.

.. _core-distributing-addins:

Distributing Addins
-------------------

You can distribute your own package of code and plugins by renaming a copy of `xloil.xll`
and creating an assoicated `ini` file.  xlOil still needs to find the core and plugin dlls, 
so you can do one of:

1) Include them with your xll
2) Ensure the main `xloil.xll` is registerd as an Excel addin.
3) Add an ``[Environment]`` block to your ini file, adding the location of the dlls to
   the `%PATH%` enviroment variable.

For example suppose you create the following files in the same directory:

    Copy `xloil.xll` to ``myfuncs.xll``

Create a ``myfuncs.ini`` file:

::

    Plugins = ["xlOil_Python"]

    [xlOil_Python]

    LoadModules=["mypyfuncs"]

Create a file ``mypyfuncs.py``:

::

    import xloil
    @xloil.func
    def greet(who):
        return "Hello " + who

Now you can load ``myfuncs.xll`` in Excel and call the `greet` function.


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

xlOil currently supports all but RTD and Cluster functions.

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

Cached Objects
--------------

xlOil has an internal store for Excel values, which is a convenient way of 
passing arrays around a sheet and as arguments to other xlOil functions.

The function ``=xloRef(A1:B2)`` returns a cache string of the form:
``<UniqueChar>[WorkbookName]SheetName!CellRef,#``

This string can then be passed instead of the source range. The data can be 
recovered using ``=xloVal(<CacheString>)``

An example use case is where you would otherwise use a named range.

**Problem**: You have large set of data on `Sheet1` which is processed in several other 
sheets and you want to ensure that when data is added to the set, all 
functions that reference are updated.

**Solution**:

- You are disciplined and only add rows to the middle, then carefully 
  cut / paste.
- You create a named range pointing at the data and manually update it in the 
  GUI when you add data.
- You use `xloRef` on the data, extending the target range far beyond these
  existing data. xlOil will automatically trim the range back to the last
  non-blank row as it reads it.  All dependent functions can use `xloVal`
  to retrieve the data.

However, there is a disadvantage to using `xloRef`: the cache is cleared when
a workbook is closed, but Excel does not know to recalculate the `xloRef` 
functions when the workbook is reopened. Hence you need to force a sheet
recalculation using *Ctrl-Alt-F9*.
===========
xlOil Python
===========

.. toctree::
	:maxdepth: 4
	:caption: Contents:
	
	xlOil_Python_Example
	xlOil_Python_Module
	
.. contents::
	:local:

The Python plugin for xlOil allows creation of Excel functions and macros backed by Python
code.

The plugin can load a specified list of module names, adding functions to Excel's global
name scope, like an Excel addin.  The plugin can also look for modules of the form
<workbook_name>.py and load these too.  Any module which contains Excel functions is 
watched for file modifications so code changes are reflected immediately in Excel.

Have a look at `<root>/test/PythonTest.py` and the corresponding Excel sheet for lots 
of examples. 


Getting Started
---------------
Choose the plugin version corresponding to Python version you want to use, e.g. 
`xlOil_Python37.dll` for Python 3.7.  Ensure it is loaded by an entry in the main 
`xlOil.ini` file.

You may need to edit the `xlOil_Python37.ini` file to set the correct Python paths.
By default, xlOil finds the PythonCore version set it the Windows registry.

The ini file also contains a list of Python modules to import - the PythonPath is 
searched in the usual way.

The log file
----------------------
If a function doesn't appear or behave as expected, check the log file created by default
in the same directory as xlOil.xll.

A common problem is that the COM interface misbehaves either failing on start-up or failing
because of an open dialog box in Excel.  For a start-up fail, unload and reload the addin. 
For other errors try to close dialog boxes or panes and if that fails, restart Excel.

Excel Functions (UDFs)
--------------------------------
Excel supports several classes of user-defined functions:

- Macros: run at user request, have write access to workbook
- Worksheet functions: run by Excel's calculation cycle. Several sub-types:
  - Vanilla
  - Thread-safe: can be run concurrently (not very useful for Python)
  - Macro-type: can read from sheet addresses and invoke a wider variety of Excel interface functions
  - Async: can run asynchronously during the calc cycle, but not in the background
  - RTD: (real time data) background threads which push data onto the sheet when it becomes available
  - Cluster: can be packaged to run on a remote Excel compute cluster

xlOil supports all but RTD and Cluster functions.

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
------------------------
If xlOil cannot convert a returned python object to Excel, it will place it in an object
cache and return a cache reference string of the form
``UniqueChar[WorkbookName]SheetName!CellRef,#``
If a string of this kind if encountered when reading function arguments, xlOil tries to 
fetch the corresponding python object. With this mechanism you can pass python objects 
opaquely between functions. 

xlOil core also implements a cache for Excel values, which is mostly useful for passing 
arrays. The function ``=xloRef(A1:B2)`` returns a cache string similar to the one used
for Python objects. These strings are automatically looked up when parsing function 
arguments.

Local Functions
-------------------------
When loading functions from an python module associated to a workbook, i.e workbook.py
xlOil defaults to registering any declared function as "local". This means it creates a
VBA stub to invoke them so that the scope of their name is local to the workbook.

Local functions have some limitations compared to global scope ones:
- Max 28 arguments
- No async or threadsafe
- Slower due to the VBA redirect
- Workbook must be saved as macro enabled (xlsm extension)

You can override the local scope on a per-function basis.

   

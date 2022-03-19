=========================
xlOil Python Concepts
=========================

.. contents::
    :local:


Workbook Modules
----------------

When an Excel workbook is opened, xlOil tries to load the module `<workbook_name>.py` 
(this is configurable).  When registering functions from workbook modules, xlOil defaults 
to making any declared functions :ref:`xlOil_Python/Functions:Local Functions`

The function :any:`xloil.linked_workbook` when called from a workbook module retrieves 
the associated workbook path.

Another way to package python code for distribution is to create an XLL, see
:ref:`core-distributing-addins`


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
        return np.sum(coeffs * x[:,None] ** range(coeffs.T.shape[0]), axis=1)

Specifing that we expect an array argument and the data type of that array avoids the
overhead of xlOil determining the type.  There is a problem with this function:
what happens if ``x`` is two-dimensional?  To avoid this possibility we can specify:

::

    @xlo.func
    def poly(x: xlo.Array(dims=1), coeffs: xlo.Array(float)):
        return np.sum(coeffs * X[:,None] ** range(coeffs.T.shape[0]), axis=1)


Events
------

Events request a callback on various user interactions. If you are familiar  
with VBA, you may have used Excel's event model already.  Most of the workbook events 
described in `Excel.Appliction <https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)#events>`_
are available in xlOil. See the xloil.Event documention for the complete list.

Some events take reference parameters, which do not exist in python. For example, setting 
the `cancel` bool in `WorkbookBeforeSave` cancels the event.  In xlOil you need to set this
value using `cancel.value=True`.

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

This happens automatically when a module is imported or reloaded as xlOil
hooks python's import mechanism.  

If you load a module outside the normal ``import`` mechanism, you can tell 
xlOil to look for functions to register with ``xloil.scan_module(module)``. 


xloPyLoad: import and scan a python module (worksheet function)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

.. function:: xloPyLoad(ModuleName)

    Imports the specifed python module and registers any it for xloil 
    functions it contains.  Leaving the argument blank loads or reloads the
    workbook module for the calling sheet, i.e. the file `WorkbookName.py`.



Multiple addins and event loops
-------------------------------

*xlOil_Python* can be used by multiple add-ins, that is, more than one XLL
loader with its own settings and python codebase can exist in the same Excel
session.  

   * Each add-in / XLL is loaded in a background thread equipped with an `asyncio`  
     event loop.  Get the loop using :any:`xloil.get_event_loop`.
   * You can find the addin associated with the currently running code with 
     :any:`xloil.source_addin` .
   * All add-ins share the same python interpreter
   * All add-ins share the python object cache
   * Worksheet functions are executed in Excel's main thread or one of its 
     worker threads for thread safe functions
   * Async / RTD worksheet functions are executed in a dedicated xlOil Core
     event loop which you can access with ``xloil.get_async_loop()``
   * You can ask xlOil to create a separate thread & event loop for an addin.     

Although CPython supports subinterpreters, most C-based extensions, particularly
*numpy* do not, so there are no plans to add subinterpreter support at this stage.

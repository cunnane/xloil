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

Events allow for a callback on user interaction. If you are familiar with VBA, you may have used 
Excel's event model already.  Most of the workbook events described in 
`Excel.Appliction <https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)#events>`_
are available in xlOil. 

See :ref:`xlOil_Python/ModuleReference:Events` for more details on python events and :doc:`Events`
for a description of the available Excel events.

Excel events do not use return values.  However, some events take reference parameters. 
For example, `WorkbookBeforeSave` has a boolean `cancel` parameter. Setting this to True cancels the 
save.  As references to primitive types aren't supported in python, in xlOil you need to set this 
value using `cancel.value=True`.

Event handlers are (currently) global to the Excel instance, so you may need to filter by workbook name 
when handling events even if you have hooked the event in a local workbook module.

Examples
~~~~~~~~

::

    def greet(workbook, worksheet):
        xlo.Range(f"[{workbook}]{worksheet}!A1") = "Hello!"

    xlo.event.WorkbookNewSheet += greet


Registering functions in other modules
--------------------------------------

xlOil automatically scans modules when they are imported or reloaded via a
hook in python's import mechanism.  This ensures any :any:`xloil.func` 
decorated functions are registered. 

If you load a module outside the normal ``import`` mechanism, you can tell 
xlOil to look for functions to register with :any:`xloil.scan_module`. 

Also see :any:`xlOil_Python/Functions:Dynamic Registration`, which explains
how any python callable can be registered as an Excel function.


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

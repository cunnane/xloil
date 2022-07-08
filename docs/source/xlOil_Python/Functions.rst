==================================
xlOil Python Registering Functions
==================================

.. contents::
    :local:

There are several ways a python function can be registered with Excel via arguments to 
:any:`xloil.func` decorator.

::

    @xloil.func
    def Greeting(who):
        return "Hello  " + who


Local Functions
---------------

When registering functions from :ref:`xlOil_Python/Concepts:Workbook Modules`, xlOil defaults to making
any declared functions "local": this means their scope is limited to the workbook.
It also means the function is automatically macro-type (xlOil achieves this by creating 
a VBA stub to invoke them).

This behaviour can be overriden by `local` argument:

::

    @xloil.func(local=False)
    def Greeting(who):
        return "Hello  " + who


Local functions have some limitations compared to global scope ones:
- No native async or threadsafe, but RTD async is OK
- Slower due to the VBA redirect
- Associated workbook must be saved as macro-enabled (xlsm extension)
- No function wizard help, but CTRL+SHIFT+A to show argument names is available

(Technical note: It is possible to use the Application.MacroOptions call to add help to the 
function wizard for VBA, but identically named functions will conflict which rather defeats 
the purpose of local functions).


Async and RTD Functions
-----------------------

RTD (real time data) functions are able to return values independently of Excel's 
calculation cycle and correspond to `async generators <https://www.python.org/dev/peps/pep-0525/>`_
in python.  For example, the function below returns the time every two seconds:

::

    import xloil, datetime, asyncio

    @xloil.func
    async def pyClock():
        while True:
            await asyncio.sleep(2)
            yield datetime.datetime.now()

This is discussed in detail in :ref:`xlOil_Python/Rtd:Introduction`.


Commands, Macros & Subroutines
------------------------------

'Macros' in VBA are declared as subroutines (``Sub``/``End Sub``) and do not return a value. 
These functions run outside the calculation cycle, triggered by some user interaction such
as a button.  They run on Excel's main thread and have full permissions on the Excel object 
model.  In the XLL interface, these are called 'commands' in the XLL interface and xlOil uses 
this terminology.

Programs which heavily use the :ref:`xlOil_Python/ExcelApplication:Introduction` object model are usually written as 
commands.

::

    @xloil.func(command=True)
    def pressRunTests():

        r = xloil.Range("TestArea")
        r.clear()
        r.set("Foo")

        ...

If not :ref:`xlOil_Python/Functions:Local Functions`, XLL commands are hidden and not displayed in 
dialog boxes for running macros, such as Excel's macro viewer (Alt+F8). However their 
names can be entered anywhere a valid command name is required, including in the macro
viewer.


Multi-threaded functions
------------------------

Declaring a function re-entrant tells Excel it can be called on all of its calculation
threads simultaneously - any other function is invoked on the main thread.  

:ref:`xlOil_Python/Functions:Local Functions` cannot be declared re-entrant.

Since python (at least CPython) is single-threaded there is no direct performance
benefit from enabling this. However, if you make frequent calls to C-based libraries 
speed gains may be possible.

::

    import xloil, ctypes

    @xloil.func(local=False, threaded=True)
    def threadsafe(x: float) -> int:
        # Do lots of calculations
        ...
        # Return the thread ID to prove the functions were executed on different threads
        return ctypes.windll.kernel32.GetCurrentThreadId(None)

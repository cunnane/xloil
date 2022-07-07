=========================
xlOil Jupyter Interaction
=========================

.. contents::
    :local:
    
Introduction
------------

xlOil can connect to a Jupyter python kernel, allowing interaction between a Jupyter notebook 
and Excel.  To use this functionality, either run `=xloPyLoad("xloil.jupyter")` in a cell or
load the `xloil.jupyter` module by specifying it in the `xlOil.ini` file:

::

    ...

    [xlOil_Python]
    LoadModules=["xloil.jupyter"]

To establish the connection, call the `xloJpyConnect` function. You pass it one of the following:

   1. The name of a notebook e.g. `MyBook.ipynb`. In this case all local jupyter instances
      will be searched to find which one has this notebook open.
   2. The full URL to the notebook, e.g. 
      `http://localhost:8888/notebooks/MyBook.ipynb?token=ac3894ab667fa1f3e4f7fe473fa89566a1580cdb49a2649b`
   3. The `kernel-xxxx-xxx-xxx.json` file which is specified in the output of running 
      the magic `%connect_info` in a Jupyter cell (no file path is required).

The `xloJpyConnect` function will return a cache reference.

.. note:: 

    The targeted jupyter kernel can be running any version of Python 3 and does not need to
    to have the `xloil` package installed.


Registering an Excel function from Jupyter
------------------------------------------

After the connection is estabilished, any function created in the kernel and decorated with
`@xloil.func` will be registed in Excel just as if it had been loaded by xlOil in the normal way.
The function will be run in the context of the kernel and the result returned asynchronously 
to Excel.

.. note:: 

    Any `@xloil.func` declarations prior to connection will be ignored, so part of the 
    jupyter notebook may need to be re-calculated to ensure functions are registed


Watching a variable in a Jupyter notebook
-----------------------------------------

The function `=xloJpyWatch(Connection, VarName)` can dynamically capture the value of any 
global variable in the kernel.  It is an RTD function so automatically updates when the variable
is changed.

Running code in the kernel
--------------------------

Calling `xloJpyRun` executes the provided string as python code in the kernel, captures the 
result and returns it to Excel.  xloJpyRun processes the code string as a `format` string using  
passing the *repr* of any additional arguments, that is, it executes 
`code_string.format(repr(arg1), repr(arg2), ...)`

Examples
--------

In an Excel cell enter:

::

    =xloJpyRun(<connection>, "{} + {}", 3, 4)


Execute the following in a jupyter cell in a connected notebook:

::

    @xloil.func
    def jptest(x):
        return f"Jupyter says {x}"

When xlOil connects to the kernel it will automatically import xlOil, although it does 
not cause a problem if re-imported.

Now try entering `=jptest("hi")` in Excel!


Using COM Automation
--------------------

If the *jupyter* kernel has the *xloil* package installed, we can turn the tables on the 
connected Excel application and control it using COM automtion. Executing the following
in the jupyter kernel will add a new workbook to the connected Excel:

::

    app = xloil.app()
    app.workbooks.add()

See :ref:`xlOil_Python/ExcelApplication:Introduction` for full details on the :any:`xloil.Application` 
object and COM automation support.


Limitations
-----------

Functions declared in the kernel cannot specify the `async` or `rtd` arguments: they are 
automatically of RTD async type to stop the kernel blocking Excel's calculation cycle!  
They cannot be multi-threaded, although xlOil can connect to more than one kernel 
simultaneously and exection in each kernel will be concurrent.

Registered kernel-based functions have addin/global scope in Excel.  Any `@xloil.func` 
declarations prior to connection will be ignored, even if `xloil` was imported. 

There is reasonable overhead in the machinery required to pass the function arguments to 
jupyter and process the result: all transport is via strings, so peformance degredation 
may be noticable for a large number of calls or for large arrays.


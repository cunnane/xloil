=========================
xlOil Jupyter Interaction
=========================

.. contents::
    :local:
    
Introduction
------------

xlOil can connect to a Jupyter python kernel, allowing interaction between a Jupyter notebook 
and Excel.  To use this functionality, either run `=xloPyLoad("xloil.jupyter")` in a cell or
load the `xloil_jupyter` module by specifying it in the xlOil ini file like this:

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

Registering an Excel function from Jupyter
------------------------------------------

After the connection is estabilished, any function created in the kernel and decorated with
`@xloil.func` will be registed in Excel just as if it had been loaded by xlOil in the normal way.
The function will be run in the context of the kernel and the result returned asynchronously 
to Excel.


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


Limitations
-----------

Functions declared in the kernel cannot use the `async` or `rtd` arguments: they are already
implictly both!  They also cannot be multithreaded, although xlOil can connect to more than 
one kernel simultaneously and exection in each will be concurrent.

The kernel-based functions are registered at global scope, so name collisions may occur as
with loading any non-local python module.

There is reasonable overhead in the machinery required to pass the function arguments to 
jupyter and process the result: all transport is via strings, so peformance degredation 
may be noticable for a large number of calls.


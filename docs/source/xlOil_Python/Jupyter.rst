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

Now any function created in the kernel and decorated with `xloil.func` will be registed in Excel
just as if it had been created in the normal way in a python file.  The function will be run 
in the context of the kernel and the result returned asynchronously to Excel.

In addition, the function `xloJpyWatch` can dynamically capture the value of any global variable
in the kernel.

Examples
--------

After following the steps in the introduction, execute the following in a jupyter cell:

::

    @xloil.func
    def jptest(x):
        return f"Jupyter says {x}"

When xlOil connects to the kernel it will automatically import xlOil, although we can do 
this manually if the python package is installed.

Now try entering `=jptest("hi")` in Excel!


Limitations
-----------

Functions declared in the kernel cannot use the `async` or `rtd` arguments: they are already
implictly both!  They also cannot be multithreaded, although xlOil can connect to more than 
one kernel simultaneously and exection in each will be concurrent.

The kernel-based functions are registered at global scope, so name collisions may occur.

There is considerable overhead in the machinery required to pass the function arguments to jupyter
and process the result.


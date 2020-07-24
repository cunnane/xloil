=========================
xlOil Jupyter Interaction
=========================

Introduction
------------

xlOil has the ability to connect to a Jupyter python kernel, allowing interaction between a 
Jupyter notebook and Excel.  To use this functionality, load the `xloil_jupyter` module by 
specifying it in the xlOil ini file like this:

::

    ...

    [xlOil_Python]
    LoadModules=["xloil_jupyter"]

To establish the connection, run the magic `%connect_info` in a Jupyter cell.  The output should
look like:

::

    ...
    or, if you are local, you can connect with just:
        $> jupyter <app> --existing kernel-ae871f5f-344a-4f63-9277-60ae72032bd7.json
    ...

You need to pass the name of the json file (no file path is required) into the `xloJpyConnect`
function, which will return a cache reference.

Now any function created in the kernel and decorated with `xloil.func` will be registed in Excel
just as if it had been created in the normal way in a python file.  The function will be run 
in the context of the kernel and the result reutrned asynchronously to Excel.

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


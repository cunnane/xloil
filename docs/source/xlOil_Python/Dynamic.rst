========================================
xlOil Python Dynamic Runtime Interaction
========================================

xlOil function registration can be controlled dynamically at runtime.  Also,
the Excel can be controlled from macros or the console.


.. contents::
    :local:

Dynamic Import
--------------

Functions for registration can be specified at runtime without the need to 
add an :any:`xloil.func` decorator.

.. note::
    Although Excel will let you, avoid doing this from (non-async) worksheet functions
    since creating new functions *during* Excel's calculation cycle is likely to cause
    instability.

The :any:`xloil.import_functions` call provides an analogue of ``from X import Y as Z`` 
with Excel UDFs.  A simple usage is:

::

    xloil.import_functions("c:/lib/AnotherFile.py", names=["greet"], as_names=["PyGreet"])


where AnotherFile.py contains:

::

    def greet(x:str):
        return f"Hello {x}!"

We specify the Excel name of the function explicitly, if we omitted this, the function 
would be registered with its python name.  In Excel you can then use the formula 
``=greet("World")``.

Typing annotations are respected, as are doc-strings - the import behaves as if we had 
decorated the function with :any:`xloil.func`.

In a worksheet, :any:`xloil.import_functions` is exposed as ``xloImport`` with the same 
arguments.

Since the import machinery can register *any* callable, including class constructors,
you cane be a little creative.  For example, the following cell formulae will
create a *pandas* *DataFrame* from the range `C1:F5`, sum over rows and take the average
of the result.

::

    [A1] : =xloImport("pandas","DataFrame")

    [A2] : =DataFrame(C1:F5)

    [A3] : =xloAttr(xloAttrObj(A2,"sum",{"axis",1}), "mean")


Notice we used ``xloAttrObj`` - the output of this is always a cache reference.  This stops 
xlOil from trying to convert the result to an Excel value.  Since a *DataFrame* is iterable
it would otherwise output *DataFrame.index* as an array.  Also note the convenient use of
`array constants <https://support.microsoft.com/en-us/office/use-array-constants-in-array-formulas-477443ea-5e71-4242-877d-fcae47454eb8>`_
to specify keyword arguments.


Dynamic Registration
--------------------

Functions can be registed using :any:`xloil.register_functions` and deregistered
with :any:`xloil.deregister_functions`.

For example:

::

    def Greet(x):
        return f"Hello {x}"
    
    xlo.register_functions([Greet])

Any callable can be registered, for example:

::

    class Closure:
        self._total = 0
        def __call__(self, x):
            self._total += x
            return x
    
    xlo.register_functions(Closure())

The name and help of the function can be controlled using :any:`xloil.func`
and the function can be linked to a specific python module, which means
it will be removed if the module is unloaded or reloaded.

::

    xlo.register_functions(
        [xlo.func(fn=Closure(), name=f"Dynamic1", register=False)], 
        module=sys.modules[__name__])

Functions are deregistered by name:

::

    xlo.deregister_functions("Greet")

==================================
xlOil Built-in Worksheet Functions
==================================

The the important ``xloRef`` and ``xloVal`` functions are described at
:ref:`core-cached-objects`

xloPyLoad: import and scan a python module (worksheet function)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

.. function:: xloImport(ModuleName:str, From=None, As=None)

    Loads functions from a module and registers them in Excel. The functions do not have to be decorated.
    Provides an analogue of `from X import Y as Z`

    ModuleName:
        A module name or a full path name to a target py file. If empty, the workbook module
        with the same name as the calling workbook is (re)loaded.
    From: 
        If omitted, imports the specified module as normal, i.e. :ref:`xloil.scan_module.` If a 
        value or array, registers only the specified object names.  If "*", all objects are registered.
    To: 
        Optionally specifies the Excel function names to register in the same order as `From`.
        Should have the same length as `From`.

    See :ref:`xloil.import_functions`

.. function:: xloAttr(Object, Name:str, *Args, **Kwargs)
    Returns the named attribute value, or the result of calling it if possible. ie, ``object.attr`` 
    or ``object.attr(*args, *kwargs)`` if it is a callable.

    Object: 
        The target object
    Name: 
        The name of the attribute to be returned.  The attribute can be a bound method,
        member, property, method, function or class
    Args
        If the attribute is callable, it will be called using these positional arguments
    Kwargs
        If the attribute is callable, it will be called using these keyword arguments
   
.. function:: xloAttrObj(Object, Name:str, *Args, **Kwargs)
    Returns the value of named attribute or the result of calling the attribute.  Behaves like
    `xloAttr` except always returns a cache object (see :ref:`xlOil_Python/TypeConversion:Cached Objects`).

    This function is useful to stop the default conversion to Excel, for example when returning
    an iterable. A typical use is when chaining `=xloAttrObj` calls.

.. function:: xloPyDebug(Debugger)

    See :ref:`xlOil_Python/Debugging:Selecting the debugger programmatically`
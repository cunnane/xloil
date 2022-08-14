==============================
xlOil Python Debugging
==============================

Visual Studio
-------------
Visual Studio Pro with Python Tools installed can break into xlOil python code.  Attach to the
relevant Excel process selecting both Python & Native debugging.

Exception Debugging
-------------------
xlOil can be configured to break into a debugger when an exception occurs in user code.  To 
do this execute the following in a loaded module:

::

    import xloil.debug
    xloil.debug.exception_debug('pdb')

Alternatively, excecute `=xloPyDebug("pdb")` in a cell; give no argument to turn off debugging.



Current debuggers supported are:

    * 'pdb': opens a console window with pdb active at the exception point
    * None: Turns off exception debugging

.. note:
    It used to be possible to select the 'vs' debugger and use Python Tools for Visual Studio 
    to open a connection on `localhost:5678`, but this package has been deprecated and no 
    longer appears to work as expected.

If `exception_debug` is specified more than once, the last value is used. It is a global but
not persistent setting.

==============================
xlOil Python Debugging
==============================

Visual Studio
-------------
*Visual Studio 2019* with Python Tools installed can break into xlOil python code.  Attach to the
relevant Excel process by selecting *only Python Code* debugging.  As of Aug 2022, mixed mode debugging
(both Python and Native C code) does not work - Python breakpoints are not hit.

VS Code
-------

*VS Code* can break into xlOil python code. You need to start the debug server in your code with

::

    import debugpy
    debugpy.listen(('localhost', 5678))

Then use the python extension's remote debugger to connect.

Pdb
---
Follow instructions for Exception Debugging, the use the command `breakpoint()` to trigger
the debugger


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
    for exception debugging but this no longer appears to work as expected.

If `exception_debug` is specified more than once, the last value is used. It is a global but
not persistent setting.

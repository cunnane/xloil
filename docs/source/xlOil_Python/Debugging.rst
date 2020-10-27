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
    * 'vs': uses ptvsd (Python Tools for Visual Studio) to enable Visual Studio or VS Code 
      to connect via a remote session. Connection is on the default settings
      i.e. localhost:5678. This means your `lauch.json` in VS Code should be:
    
        ::

            {
                "name": "Attach (Local)",
                "type": "python",
                "request": "attach",
                "localRoot": "${workspaceRoot}",
                "port": 5678,
                "host":"localhost"
            }

      A breakpoint is also set a the exception site.
    * None: Turns off exception debugging

If `exception_debug` is specified more than once, the last value is used. It is a global but
not persistent setting.

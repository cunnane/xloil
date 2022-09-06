==============================
xlOil Python Debugging
==============================

Only python debuggers capable of handling *embedded* python interpreters can be used to debug xlOil
python code.

Visual Studio
-------------

*Visual Studio 2019* and *Visual Studio 2022* with Python Tools installed can break into xlOil python code. 
Attach to the relevant Excel process selecting *only Python Code* debugging.  

Visual Studio will occasionally hang when first hitting a python breakpoint. In the case, restart VS and
Excel and re-try

.. note::
    As of Aug 2022, mixed mode debugging with both Python and Native C code does not work: Python 
    breakpoints are not hit.


VS Code
-------

*VS Code* can break into xlOil python code. You need to tell xlOil to start a `debugpy` server by
selecting it as the debugger in the xlOil ribbon or the ini file. Then use VS Code's python 
remote debugger to connect.  The default port is *5678* but this can be changed in the ini file.

*VS Code* cannot hit breakpoints in async and RTD functions even though tracing is enabled in the
relevant threads.  The reason for this is unknown.

.. note::
    Running the *debugpy* server has some peformance implications, so avoid leaving debugging 
    enabled when not required.

Pdb
---

Pdb can be used for post-mortem debugging, i.e. after an exception occurs. Select *pdb* in as 
the debugger in the xlOil ribbon or the ini file.  Use the command `breakpoint()` in your 
code to trigger the debugger, or wait for an exception.


Selecting the debugger programmatically
---------------------------------------

A debugger can be selected at runtime in the xlOil ribbon but this choice will persist when Excel
is restarted.  Alternatively, the debugger can be choosen in code, which is not persistent. Use
the following python code:

::

    import xloil.debug
    xloil.debug.use_debugger('pdb')

Or the Excel formula:

::

    =xloPyDebug(...)


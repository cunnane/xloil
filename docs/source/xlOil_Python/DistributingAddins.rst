================================
xlOil Python Distributing Addins
================================

We might like to distribute our code as a packaged addin so users don't have to edit
their `xlOil.ini` to load it.  To do this, run the following at a command prompt:

::

    xloil create myaddin.xll

This will create a `myaddin.xll` and `myaddin.ini` in the current directory.
By default, the XLL will try to load `myaddin.py` in the same directory, so let's 
create it:

::

    import xloil as xlo

    @xlo.func
    def MySum(x, y, z):
        '''Adds up numbers'''
        return x + y + z

Now drop `myaddin.xll` into an Excel session and try to use ``MySum``.

Users of `myaddin.xll` will need access to a python distribution which contains
the *xlOil* package, however they do not need the xlOil addin installed.

You may be able to place the python distribution on a shared drive and modify
the `PYTHONPATH`, `PYTHONHOME` and `PATH` environment variables `myaddin.ini`
to point to it.

.. important:: 
    If a user of `myaddin.xll` has an ini file at `%APPDATA%\xlOil\xlOil.ini``
    the core xlOil.dll is loaded using those settings before `myaddin.xll`.
    The assumption is that the user has the xlOil addin installed in Excel, but 
    since only one instance of xlOil (and one python interpreter) can be hosted in 
    Excel, one settings file must take precedence. You can make your addin take
    precedence with the `LoadBeforeCore` flag in `myaddin.ini`.

Installing the Addin
--------------------

If you copy `myaddin.xll` and `myaddin.ini` to your `%APPDATA%\Microsoft\Excel\XLSTART` 
directory, you'll notice that Excel opens the ini file, which is not ideal! xlOil will
look for `myaddin.ini` in `%APPDATA%\xlOil`, so install the ini to there. You'll also need 
to choose a directory to hold your python modules and ensure `myaddin.ini` points to it:
`%APPDATA%\xlOil\myaddin` would be a sensible choice.

For more on packaging addins, see :ref:`core-distributing-addins`.
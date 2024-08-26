=======================================
xlOil Python Distribution and Packaging
=======================================

xlOil supports two possiblilities for distribution:
   * Creating an XLL addin to distribute code to existing xlOil users
   * Packaging Python and xlOil to install code for new users


Creating an Addin
=================

The idea of this approach is to create an XLL addin so users don't have to edit
their `xlOil.ini` to load some packaged code.  

Users of XLL will need access to a python distribution which contains
the *xlOil* package, however they do not need the xlOil addin installed.

At a command prompt, run:

::

    xloil create myaddin.xll

This will create a `myaddin.xll` and `myaddin.ini` in the current directory.
By default, the XLL will try to load `myaddin.py` in the same directory, so 
create this file:

::

    import xloil as xlo

    @xlo.func
    def MySum(x, y, z):
        '''Adds up numbers'''
        return x + y + z

Dropping `myaddin.xll` into an Excel session will create the function ``MySum``.

If the python distribution will be in a standard directory on users' machines,
the `PYTHONEXECUTABLE` and `PATH` environment variables in `myaddin.ini` can
be used to point to it, otherwise these variables can read the python location
from the registry.

If you copy `myaddin.xll`, `myaddin.ini` and `myaddin.py` to your `%APPDATA%\\Microsoft\\Excel\\XLSTART` 
directory, Excel attempts to opens the ini file and py files which is not ideal! xlOil will
look for `myaddin.ini` in `%APPDATA%\\xlOil`, so install the ini to there instead. You'll also need 
to choose a directory to hold `myaddin.py` (or other python modules) and ensure `myaddin.ini` points to 
it; `%APPDATA%\\xlOil\\myaddin` could be a sensible choice.

.. important:: 
    If a user of `myaddin.xll` has an ini file at `%APPDATA%\\xlOil\\xlOil.ini``
    the core xlOil.dll is loaded using those settings before `myaddin.xll`.
    The assumption is that the user has the xlOil addin installed in Excel, but 
    since only one instance of xlOil (and one python interpreter) can be hosted in 
    Excel, one settings file must take precedence. You can make your addin take
    precedence with the `LoadBeforeCore` flag in `myaddin.ini`.


Packaging Python
================

xlOil can use `PyInstaller <https://pyinstaller.org/>`_ to package a python distribution and
create an installer executable.  Support for this is fairly rudimentary at present.

You shouuld start with a minimum python distribution (ideally based on the standard distribution)
for the code you want to distribute, otherwise *PyInstaller* may create a very large output.  The
distribution should include xlOil and be able to load your code in Excel.

To package the settings in the file *myaddin.ini* which loads the python module *excel_funcs.py*, 
run the following command:

::

    xloil package myaddin.ini --hidden-import excel_funcs

Note that *--hidden-import* is actually an `argument to PyInstaller <https://pyinstaller.org/en/stable/usage.html#options>`_
and can be specified multiple times.  Any other trailing arguments will be passed directly to *PyInstaller*.

The resulting *dist* directory will contain:

  * install_main.exe (installs xlOil)
  * _internal (contains the python distribution)

The installer does not copy the python distribution, it is used in-situ


Customising the packaging
-------------------------

Calling 

:: 

    xloil package -makespec myaddin.ini
    
stops the packaging process before after creation of the 
`PyInstaller spec files <https://pyinstaller.org/en/stable/spec-files.html>`_.  You can edit this
spec file directly as described in the *PyInstaller* docs, then invoke *PyInstaller* on the spec file
yourself to finish the process.

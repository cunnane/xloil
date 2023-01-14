===========
xlOil Core
===========

Here we tell you how to get started using xlOil and how to distribute addins

.. _core-getting-started:

Getting Started
---------------

You need **Excel 2010** or later on a desktop PC. xlOil will not work with online 
or Mac versions of Office.

.. important::

    If you want to use xlOil with python you should install it via `pip`. Stop reading
    this page and go to :doc:`xlOil_Python/GettingStarted`

Download the binary package (e.g. from gitlab) and unzip to a directory of 
your choice. 

You can run the `xlOil_Install.ps1` script to install the addin for every
Excel session, or just drop `xlOil.xll` into any running Excel session
to load xlOil temporarily.

xlOil should now load when you open Excel, try following 
:any:`sql-getting-started`

To configure the plugins being loaded, see :any:`core-edit-settings-files`.
Some plugings, such as the python one, have several paths which must be set 
correctly - it will generally be easier to use these plugins by following their
specific installation instructions.

.. _core-example-sheets:

Example sheets
--------------

To check your setup and see some of the capabilities of xlOil, try:
:download:`Tests and Examples </_build/xlOilExamples.zip>`.

.. _core-edit-settings-files:

Editing the settings file
-------------------------

There is an `xlOil.ini` file linked to the main `xlOil.xll` addin. (This ini file 
is actually parsed as TOML, an extension of the ini format). xlOil searches for
this file first in `%APPDATA%/xlOil` then in the directory containing the `xlOil.xll` 
addin. 

The two most important setting in `xlOil.ini` are:

::

    ...
    XLOIL_PATH='''C:\lib\xloil```
    ...
    Plugins=["xloil_Python.dll", "xlOil_SQL.dll"]

``XLOIL_PATH`` allows the `xlOil.xll` addin to locate the main xlOil DLL if the 
addin is being run from a different directory.  When the main DLL has loaded, 
xlOil loads the specified plugins. It searches for these plugins first in the 
directory containing `xlOil.dll`, then the directory containing the XLL, then 
the usual DLL search paths. 


Setting enviroment variables in settings files
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Each plugin can have an *Environment* section in the settings file. Within this block
tokens are interpreted as enviroment variables to set. A plugin's environment settings 
are processed before the plugin is loaded. 

Keys are interpreted as environement variables to set. Values can reference other enviroment 
variables by surrounding the name with `%` characters.

In addition you can pull values from the registry by surrounding the registry
path with angle brackets `<>`. Leaving a trailing backslash `\\` in the 
registry path fetches the default value for that key.

The TOML syntax of three single quotes indicates a string literal, this avoids escaping 
all the backslashes.

The default enviroment block for Python looks like this:

::

    [[xlOil_Python.Environment]]
    xlOilPythonVersion="3.7"

    [[xlOil_Python.Environment]]
    PYTHONPATH='''<HKLM\SOFTWARE\Python\PythonCore\%xlOilPythonVersion%\PythonPath\>'''
    PYTHON_LIB='''<HKLM\SOFTWARE\Python\PythonCore\%xlOilPythonVersion%\InstallPath\>'''

    [[xlOil_Python.Environment]]
    PATH='''%PATH%;%PYTHON_LIB%;%PYTHON_LIB%\Library\bin'''

The double brackets tell TOML that the order of these declarations is important,
this means we can refer to previously set enviroment variables.

In addition you can pull values from the registry by surrounding the registry
path with angle brackets `<>`, for example, 
`<HKLM\SOFTWARE\Python\PythonCore\3.6\InstallPath\ExecutablePath>`. 
Leaving a trailing backslash `\\` in the registry path fetches the default 
value for that key.

Troubleshooting
---------------

You may need to tweak your settings file: :any:`core-edit-settings-files`

A common problem is that the COM interface misbehaves by either failing on start-up or failing
because of an open dialog box in Excel.  For a start-up fail, unload and reload the addin. 
For other errors try to close dialog boxes or panes and if that fails, restart Excel.

The log file
~~~~~~~~~~~~~

If xlOil detects a serious load error, it pops up a log window to alert you (this can
be turned off). If it succesfully loaded the core DLL a log file will also be created
next to `xlOil.ini`, which by default is in ``%APPDATA%\xlOil``.  If xlOil loaded, the 
worksheet function `xloLog` can tell you where this file is.  A setting in `xlOil.ini` 
controls the log level.

Manual installation
~~~~~~~~~~~~~~~~~~~

The `xlOil_Install.ps1` script does the following:

   1. Check xlOil is not in Excel's disabled add-ins
   2. Copy xlOil.xll to the ``%APPDATA%\Microsoft\Excel\XLSTART`` directory
   3. Copy xlOil.ini in the ``%APPDATA%\xlOil``` directory
   4. Check VBA Object Model access is allowed in 
      `Excel > File > Options > Trust Center > Trust Center Settings > Macro Settings``


Manual removal
~~~~~~~~~~~~~~

Should you need to force remove xlOil, do the following:

   1. Remove *xlOil.xll* from ``%APPDATA%\Microsoft\Excel\XLSTART``
   2. Remove the directory ``%APPDATA%\xlOil```

If you have added *xlOil.xll* or another xll add-in (xlOil does not do this by default)
and you want to remove it go to:

   1. `Excel > File > Options > Add-ins > Manage Excel Addins`
   2. If the previous step fails to remove the addin, start Excel with elevated/admin 
      priviledges and retry
   3. If that fails, try to remove the add-in from the registry key
      ``HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\<Version>\Excel\Options``.
      You should see values *OPEN*, *OPEN1*, etc with add-in names to be loaded. After removing
      a value, you need to rename the others to preserve the numeric sequence.
   4. If that does not work, also look at this registry key:
      ``HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\<Version>\Excel\Options``.

Note you may need to run the registry editor with elevated priviledges.

To really scrub the registry, you may find references to the addin under:
   * `HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Office\\<Version>\\Excel\\Add-in Manager`
   * `HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Office\\<Version>\\Excel\\AddInLoadTimes`
   * `HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Office\\<Version>\\Excel\\Resiliency\DisabledItems`
   * `HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Office\\Excel\\Addins`
   * `HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Office\\Excel\\AddinsData`

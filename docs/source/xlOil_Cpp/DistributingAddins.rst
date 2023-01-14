=========================
Distributing xlOil Addins
=========================

You can distribute your own package of code and plugins by renaming a copy of `xloil.xll`
and creating an assoicated `ini` file. 

In C++, there is the option to compile a static XLL, independent of `xlOil.dll`, see 
:doc:`/xlOil_Cpp/StaticXLLs`.  In this case, distribution is simple: copy the XLL and 
ensure and DLLs you may have linked are on the user's PATH.

.. note::

    Python users can make use the of xlOil installer, see :doc:`/xlOil_Python/DistributingAddins`
    which will correctly set the paths in `xlOil.ini` for the distribution it is run under.

Example
-------

Suppose you have created a plugin as described in :doc:`/xlOil_Cpp/GettingStarted` called
`my_plugin.dll`.

Copy `xloil.xll` as `myaddin.xll`. Create a `myaddin.ini` file in the same directory 
which contains:

::

    Plugins = ["my_plugin"]
    
    [Addin.Environment]
    PATH='''%PATH%;c:\path\to\DLLs'''


The directory ``c:\path\to\DLLs`` must contain your `my_plugin.dll` and `xlOil.dll`. 
Alternatively you simply include all the DLLs in the same directory as `myaddin.xll`.

If you copy `myaddin.xll` to ``%APPDATA%\Microsoft\Excel\XLSTART`` to have Excel 
load it, you cannot put `myaddin.ini` or any other DLLs in the same directory, or Excel
will try to load them as well. xlOil will look for `myaddin.ini` in ``%APPDATA%\xlOil``,
so copy the ini file to that location instead.

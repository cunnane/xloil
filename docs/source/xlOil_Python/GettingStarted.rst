=========================
xlOil Python Introduction
=========================

.. contents::
    :local:

Introduction
------------

The Python plugin for xlOil allows creation of Excel functions and macros backed by Python
code.

xlOil_Python loads specified python modules, looking for functions to add to Excel's global
name scope, like an Excel addin.  The plugin can also look for modules of the form
<workbook_name>.py and load these too, this is like creating a VBA code module for a workbook.
Any python module which contains Excel functions is watched for file modifications so 
code changes are reflected immediately in Excel.

xlOil_Python is tightly integrated with numpy, allowing creation of fast Excel array 
functions.

For examples, have a look at :doc:`Example` and :ref:`core-example-sheets`.

Getting Started
---------------

**You must use the same bit-ness of python and Excel**.  So if your Excel is 32-bit, you must
install xloil using a 32-bit python distibution.

Run the following at a command prompt with python environment settings:

::

    pip install xlOil
    xloil install

The call to ``xloil install`` registers the xlOil addin with Excel and places a settings
files at the `%APPDATA%/xlOil/xlOil.ini`.  The settings file describes the python modules 
which will be loaded and sets the paths to the python distribution. xlOil attempts 
to set the python paths automatically, but can be overriden if required.

To test the setup, you can try the python example sheet: :ref:`core-example-sheets`.

Note it's not necessary for ``xlOil.xll`` to be registered in this way: you can just
drop it into your Excel session when required. 

You now have three ways to get xlOil to load your python code:


My first xlOil module
~~~~~~~~~~~~~~~~~~~~~

Now let's make our first python module using xlOil.  Create a `MyTest.py` file with 
the following lines:

::

    import xloil as xlo

    @xlo.func
    def Greeting(who):
        return "Hello  " + who

Edit `%APPDATA%/xlOil/xlOil.ini` so that `PYTHONPATH` includes the 
directory containing `MyTest.py` then add `MyTest` to the `LoadModules` key.

Now open Excel and call the Greeting function.


My first workbook module
~~~~~~~~~~~~~~~~~~~~~~~~

Create and an Excel workbook called `MyBook`. In the same directory, create 
a file `MyBook.py` containing the following:

::

    import xloil as xlo

    @xlo.func
    def Adder(x, y):
        return x + y

You'll need to open and close `MyBook` in Excel for xlOil to find the python file.
Now try invoking the Adder function - it can also add arrays!

If this isn't working, ensure that "Trust accesst to the VBA object model" 
is checked in Excel Options -> Trust Centre -> Macro Settings.


My first xlOil addin
~~~~~~~~~~~~~~~~~~~~~

We might like to distribute our code as a packaged addin so users don't have 
to edit `xlOil.ini`. To do this, run the following at a command prompt:

::

    xloil create myaddin.xll

This will create a `myaddin.xll` and `myaddin.ini` in the current directory.
By default, the XLL will try to load `myaddin.py`, so let's create it:

::

    import xloil as xlo

    def MySum(x, y, z):
        '''Adds up numbers'''
        return x + y + z

Now drop `myaddin.xll` into an Excel session and try to use ``MySum``.

For more on this packaging addins, see :ref:`core-distributing-addins`.


Troubleshooting
~~~~~~~~~~~~~~~

Check the `xlOil.log` file for errors. By default, the log file is created in the
same directory as `xlOil.ini` in your `%APPDATA%/xlOil` directory.  If xlOil core has 
succesfully loaded, the worksheet function `xloLog` will tell you where this file is.

You may need to set the python paths in the `xlOil.ini` file for xlOil to find 
your python distribution.

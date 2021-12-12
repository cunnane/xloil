=====================================
xlOil Python Excel.Application object
=====================================

The `Excel.Application` object is the root of Excel's COM interface.  If you have used VBA you 
will likely have come accross it.  In xlOil, you can get a reference to this object with 
`xloil.app()`. From there the `comtypes <https://pythonhosted.org/comtypes/>`_ or `pywin32 <http://timgolden.me.uk/pywin32-docs/html/com/win32com/HTML/QuickStartClientCom.html>`_ 
libraries provides syntax similar to VBA to call methods on the object.

The available methods are documented extensively at `Excel object model overview <https://docs.microsoft.com/en-us/visualstudio/vsto/excel-object-model-overview>`_
and `Application Object <https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)>`_

COM support can be provided by 'comtypes', a newer pure python package or 'win32com' 
a well-established more C++ based library. You can choose using:

::

    # Run this before any calls to xloil.app()
    xloil.use_com_lib("win32com") # run this to select win32com
    xloil.use_com_lib("comtypes") # the default, no need to run this explicitly



Application.Run
---------------

In VBA this takes a function name and a variable argument list and attempts to call the specified
function dynamically.  In xlOil, use :obj:`xloil.excel_func` and :obj:`xloil.excel_func_async`.

Examples
--------

We lift some examples directly from `the comtypes help <https://pythonhosted.org/comtypes/>`_

::

    xl = xlo.app()

    # Accessing a range with empty index
    X = xl.Range["A1", "C1"].Value[:]
    # X now contains a tuple like (10, "20", 31.4)

    # Alternative syntax, gives Y == X
    Y = xl.Range["A1", "C1"].Value[()]

    # Writing to a range uses the same syntax
    xl.Range["A1", "C1"].Value[:] = (3, 2, 1)
    xl.Range["A1", "C1"].Value[()] = (1, 2, 3)

    # Looks very similar but uses the xlOil range object so has slightly
    # different syntax. We're calling the *Range* constructor so we use
    # round brackets.
    xlo.Range("A1:C1").value = (1, 2, 3)


Troubleshooting
---------------

https://stackoverflow.com/questions/52889704/python-win32com-excel-com-model-started-generating-errors
=====================================
xlOil Python Excel.Application object
=====================================

The `Excel.Application` object is the root of Excel's COM interface.  If you have used VBA you 
will likely have come accross it.  In xlOil, you can get a reference to this object with 
`xloil.app()`. From there the `comtypes <https://pythonhosted.org/comtypes/>`_ or
`pywin32 <http://timgolden.me.uk/pywin32-docs/html/com/win32com/HTML/QuickStartClientCom.html>`_ 
libraries provides syntax similar to VBA to call methods on the object.

The available methods are documented extensively at `Excel object model overview <https://docs.microsoft.com/en-us/visualstudio/vsto/excel-object-model-overview>`_
and `Application Object <https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)>`_

COM support can be provided by 'comtypes', a newer pure python package or 'win32com' (also called
`pywin32`) a well-established more C++ based library. You can pass a ``lib`` argument to 
:any:`xloil.app` or the ``to_com`` method.  If omitted, the default is 'comtypes', unless changed 
in the XLL's ini file.


Commands, Macros & Subroutines
------------------------------

'Macros' in VBA are declared as subroutines (``Sub``/``End Sub``) and do not return a value. 
These functions are run outside the calculation cycle triggered by some user interaction such
as a button.  They are run on Excel's main thread and have full permissions on the Excel object 
model.  In the XLL interface, these are called 'commands' in the XLL interface and xlOil uses 
this terminology.

Programs which heavily use the `Excel.Application` object model are usually written as 
macros / commands.

Unless declared *local*, XLL commands are hidden and not displayed in dialog boxes for running 
macros, such as Excel's macro viewer (Alt+F8). However their names can be entered anywhere a 
valid command name is required, including in the macro viewer.


Calling Worksheet Functions and Application.Run
-----------------------------------------------

In VBA, ``Application.Run`` takes a function name and a variable argument list and attempts
to call the specified user-defined function.  In xlOil, use :obj:`xloil.run` to make the same 
call or go via the COM library with ``xloil.app().Run(...)``. Like all COM calls, they must be
invoked on the main thread.

To call a worksheet function, use :obj:`xloil.call`. This can also invoke old-style 
`macro sheet commands <https://docs.excel-dna.net/assets/excel-c-api-excel-4-macro-reference.pdf>`_.
It must be called from a non-local worksheet function on the main thread.  To access a worksheet
function from COM use ``xloil.app().WorksheetFunction.Sum(...)`.

:obj:`xloil.run` and :obj:`xloil.call` have async flavours :obj:`xloil.run_async` and 
:obj:`xloil.call_async` which return a future and can be called from any thread.

+------------------------+---------------------------------------------------------+-------------+
| Function               |  Use                                                    | Call from   |
+========================+=========================================================+=============+
| :obj:`xloil.run`       | Calls user-defined functions as per `Application.Run`   | Main thread |
+------------------------+---------------------------------------------------------+-------------+
| :obj:`xloil.run_async` |                                                         | Anywhere    |
+------------------------+---------------------------------------------------------+-------------+
| :obj:`xloil.call`      | Calls worksheet functions, UDFs or macro sheet commands | Non-local worksheet function |
+------------------------+---------------------------------------------------------+-------------+
| :obj:`xloil.run_async` |                                                         | Anywhere    |
+------------------------+---------------------------------------------------------+-------------+


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

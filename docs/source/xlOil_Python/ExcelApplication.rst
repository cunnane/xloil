=====================================
xlOil Python Excel.Application object
=====================================

.. contents::
    :local:

The `Excel.Application` object is the root of Excel's COM interface.  If you have used VBA you 
will likely have come across it.  In xlOil, you can get a reference to this object with 
:any:`xloil.app`. From there the `comtypes <https://pythonhosted.org/comtypes/>`_ or
`pywin32 <http://timgolden.me.uk/pywin32-docs/html/com/win32com/HTML/QuickStartClientCom.html>`_ 
libraries provides syntax similar to VBA to call methods on the object.

The available methods are documented extensively at `Excel object model overview <https://docs.microsoft.com/en-us/visualstudio/vsto/excel-object-model-overview>`_
and `Application Object <https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)>`_

COM support can be provided by 'comtypes', a newer pure python package or 'win32com'
a well-established more C++ based library.  :any:`xloil.app` or the :any:`xloil.Worksheet.to_com` 
method accept a 'lib' argument  If omitted, the default is 'comtypes'.  The default can 
be changed in the XLL's ini file.


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


Accessing Sheets and Ranges
---------------------------

xlOil mirrors a small part of the `Excel.Application` object model to provide easier
access to sheets and ranges.

We take some examples for accessing Ranges directly from 
`the comtypes help <https://pythonhosted.org/comtypes/>`_

::

    xl = xloil.app()

    # Accessing a range with empty index
    X = xl.Range["A1", "C1"].Value[:]
    # X now contains a tuple like (10, "20", 31.4)

    # Alternative syntax, gives Y == X
    Y = xl.Range["A1", "C1"].Value[()]

    # Using xlOil functions, gives Z == X
    Z = xloil.Range("A1:C1").value

    # Writing to a range uses the same syntax
    xl.Range["A1", "C1"].Value[:] = (3, 2, 1)
    xl.Range["A1", "C1"].Value[()] = (1, 2, 3)

    # Using xlOil syntax
    xloil.Range("A1:C1").value = (1, 2, 3)


Troubleshooting
---------------

Both *comtypes* and *win32com* have caches for the python code backing the Excel object model. If 
these caches somehow become corrupted, it can result in strange COM errors.  It is safe to delete 
these caches and let the library regenerate them. The caches are at:

   * *comtypes*: `...\site-packages\comtypes\gen`
   * *win32com*: run ``import win32com; print(win32com.__gen_path__)``

See `for example <https://stackoverflow.com/questions/52889704/python-win32com-excel-com-model-started-generating-errors>`_


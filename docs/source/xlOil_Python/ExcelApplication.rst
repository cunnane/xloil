==========================================
xlOil Python: The Excel.Application object
==========================================

.. contents::
    :local:

Introduction
------------

The `Excel.Application` object is the root of Excel's COM interface.  If you have used VBA you 
will likely have come across it.  If xlOil is running embedded in Excel, you can get a reference 
to the parent application with :any:`xloil.app`.  If xlOil has been imported as package, you can 
create an Application with :any:`xloil.Application`.

xlOil mirrors a small part of the `Excel.Application` object model as discussed below. For other calls,
the COM interface can be accessed directly which provides syntax similar to the `Application`` object in 
VBA.

.. important:: 
    All COM calls must be invoked on the main thread!  A function runs in the main thread if is 
    not declared multi-threaded or if it is called from VBA or the GUI. However during Excel's calc
    cycle, much of the Application object model is locked, in particular, writing to the sheet is blocked.
    To schedule a callback to be run on main thread use :any:`xloil.excel_callback`.

The object model is documented extensively at `Excel object model overview <https://docs.microsoft.com/en-us/visualstudio/vsto/excel-object-model-overview>`_
and `Application Object <https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)>`_


Calling Worksheet Functions and Application.Run
-----------------------------------------------

In VBA, ``Application.Run`` takes a function name and a variable argument list and attempts
to call the specified user-defined function.  In xlOil, use :obj:`xloil.run` to make the same 
call or go via the COM library with ``xloil.app().Run(...)``. All COM calls must be invoked
on the main thread, however :obj:`xloil.run` and :obj:`xloil.call` have async flavours 
:obj:`xloil.run_async` and :obj:`xloil.call_async` which return a Future and can be called 
from any thread.

To call a worksheet function, use :obj:`xloil.call`. This can also invoke old-style 
`macro sheet commands <https://docs.excel-dna.net/assets/excel-c-api-excel-4-macro-reference.pdf>`_.
It must be called from **a non-local worksheet function**.  Worksheet functions can be
called from COM, for example, ``xloil.app().WorksheetFunction.Sum(...)``.

+-------------------------------+---------------------------------------------------------+------------------------------+
| Function                      |  Use                                                    | Call from                    |
+===============================+=========================================================+==============================+
| :obj:`xloil.run`              | Calls user-defined functions as per `Application.Run`   | Main thread                  |
+-------------------------------+---------------------------------------------------------+------------------------------+
| :obj:`xloil.run_async`        | (as above but async)                                    | Anywhere                     |
+-------------------------------+---------------------------------------------------------+------------------------------+
| :obj:`xloil.call`             | Calls worksheet functions, UDFs or macro sheet commands | Non-local worksheet function |
+-------------------------------+---------------------------------------------------------+------------------------------+
| :obj:`xloil.run_async`        | (as above but async)                                    | Anywhere                     |
+-------------------------------+---------------------------------------------------------+------------------------------+
| xloil.app().WorksheetFunction | Calls worksheet functions                               | Main thread                  |
+-------------------------------+---------------------------------------------------------+------------------------------+

xlOil's Excel Object Model
--------------------------

xlOil mirrors a small part of the `Excel.Application` object model to faciliate easier access to the commonly 
used :obj:`xloil.Application`, :obj:`xloil.Workbook`, :obj:`xloil.Worksheet`, :obj:`xloil.ExcelWindow`, and 
:obj:`xloil.Range` objects.

Each of xlOil's application objects provides a `to_com` method which accepts an optional *lib* argument. 
Calling this returns a marshalled COM object which supports any method or property in the full Application object 
model. COM support is be provided by `comtypes <https://pythonhosted.org/comtypes/>`_ , a newer pure 
python package or `win32com <http://timgolden.me.uk/pywin32-docs/html/com/win32com/HTML/docindex.html>`_ 
a well-established C++ based library.  If omitted, the default is 'win32com'. The default can be changed 
in the XLL's ini file.

COM methods can be called directly on xlOil's application objects, so the following are equivalent:

::

    xlo.Application().RegisterXLL(...)
    xlo.Application().to_com().RegisterXLL(...)

There is no ambiguity with other methods on the *Application* object as COM methods and properties 
all start with a capital letter.

COM methods can be called with keyword arguments - note COM arguments start with a capital letter.

::

    xloil.app().Selection.PasteSpecial(Paste=xloil.constants.xlPasteFormulas)


Excel Automation
----------------

Excel's COM interface allows the application to be driven externally by a script. This is best explored
by looking at (a simplified version of) xlOil's test runner.  The test runner is started at the command line,
rather than inside an Excel instance like an xlOil-based addin.  You may want to look at the documentation
for Excel's `Name <https://docs.microsoft.com/en-us/office/vba/api/excel.name>`_ object.

::

    import xloil as xlo

    # Create a new Excel instance and make it visible
    app = xlo.Application()
    app.visible = True

    # Load addin
    if not app.RegisterXLL("xloil.xll"):
        raise Exception("xloil load failed")

    test_results = {}
    for filename in ['TestUtils.xlsx, PythonTest.xlsm']:

        # Open the workbook in readonly mode: don't change the test source!
        wb = app.open(filename, read_only=True)
    
        app.calculate(full=True)

        # Loop through all named ranges in the workbook, looking for ones 
        # prefixed with 'Test_'.  We expect those ranges to contain True
        # for a successful test outcome.
        names = wb.to_com().Names
        for named_range in names:
            if named_range.Name.lower().startswith("Test_"):
                # skip one char as RefersTo always starts with '='
                address = named_range.RefersTo[1:]
                test_results[(filename, named_range.Name)] = wb[address].value
        
        wb.close(save=False)

    app.quit()

    if not all(test_results.values()):
        print("-->FAILED<--")


Creating an Application
=======================

The :any:`xloil.Application` object can be created in several ways:

    1) When xloil is embedded, the parent applicaton object is in :any:`xloil.app()`
    2) `xlo.Application()` with no arguments opens an new instance of Excel (but does not make it visible)
    3) `xlo.Application("MyWorkbook.xlsx")` returns an instance of Excel which has *MyWorkbook.xlsx* open (or throws)
    4) `xlo.Application(ComObject)` points an Application at a COM object managed by *win32com* or *comtypes*
    5) `xlo.Application(HWND)` creates a Application given the window handle of Excel's main window as an int

The application object can be :any:`xloil.Application.quit()` manually or since it is a context manager, 
you can write:

::

    with xloil.Application() as app:
        # do stuff
        ...

    # app has been quit without saving any open Workbooks


Accessing Sheets and Ranges
---------------------------

When xlOil is embedded in Excel as an addin, there is a natural default :obj:`xloil.Application` 
object: the parent application, which can be accessed by :any:`xloil.app()`.  Additionally,
when embedded you can unambigiously create :any:`xloil.Range` and :any:`xloil.Worksheet` objects
without needing to specify the application.

Reading from a Range
====================

::

    import xloil as xlo

    # if xlOil is embedded: no need to specify Application.
    # Returns a numpy array
    xlo.Range("A1:C1").value

    # Above is equivalent to
    xlo.app().range("A1:C1").value

    # Using COM (win32com) to access a range with empty index
    # Returns a tuple rather than a numpy array
    xlo.app().Range("A1", "C1").Value

If the range referred to is empty, its `value` array will be populated with `None`. This 
is different to array/range arguments to :any:`xloil.func` worksheet functions where the
array is trimmed to the last non-blank. This behaviour can be replicated with 
:any:`xloil.Range.trim` :

::

    r = xlo.app().range("A1:C1")

    r.clear()
    r.trim().value  # returns the array [None]

    r.set(1)
    r.trim().value # returns the array [1, 1, 1]

The square bracket (getitem) operator for a Range behaves like a numpy array,
in that if the tuple specifies a single cell, it returns the value in that cell, otherwise 
it returns a :any:`xloil.Range` object.  To create a range consisting of a single cell
use :any:`xloil.Range.cells`.

Writing to a range
==================

::

    # Using the COM object
    xlo.app().Range("A1", "B2").Value = ((1, 2), (3, 4))

    rng = xlo.Range("A1:B2")
    # Using xlOil syntax (can use numpy array)
    rng.value = np.array([[1, 2], [3, 4]])

    # Set the entire range to a single value
    rng.set("hello")

    # Add something
    rng += " world!"


Using Worksheets and Workbooks
==============================

There are several ways to address or refer to part of a worksheet:

::

    wb = xloil.active_workbook()  # Only available when embedded

    # Specify external Excel range address
    r1 = xlo.app().range[f'{wb.name}Sheet1!B2:D3']

    # Specify workbook Excel range address
    r1 = wb['Sheet1!B2:D3']

    # Specify worksheet, then local Excel range address
    ws = wb['Sheet1']
    r1 = ws['B2:D3']
    
    # The range function, like in Excel includes right and left hand ends
    r2 = ws.range(from_row=1, from_col=1, to_row=2, to_col=3)

    # The slice syntax follows python conventions so only the left
    # hand end is included
    r3 = ws[1:3, 1:4]


The square bracket (getitem) operator for :ref:`xloil.Worksheet` always returns
a :ref:`xloil.Range`. For :ref:`xloil.Workbook` it may return a :ref:`xloil.Range`
or a :ref:`xloil.Worksheet`.

Writing to a worksheet
~~~~~~~~~~~~~~~~~~~~~~

::

    data = np.array([[1, 2], [3, 4]])

    ws = xloil.worksheets['Sheet1']

    # ws[...] gives a Range, so  
    ws["A1:B2"].value = data

    # However, value is optional when writing to a sheet
    ws["A1:B2"] = data  

    # You can copy another part of the sheet, it's faster to
    # drop the value property here
    ws["A1:B2"] = ws["D1:E2"] 

    # Also works for Workbooks
    wb = xloil.active_workbook()
    wb['Sheet1!B2:D3'] = ws["D1:E2"] 



Pausing Excel Calculations
--------------------------

When writing to worksheets, performance can often be improved by disabling Excel's auto calculation 
and Event model, otherwise calculation cycles and events will be triggered on each write.

This is straightforward using :any:`xloil.PauseExcel`:

::

    with xloil.PauseExcel() as paused:
        for i in range(100):
            worksheet[i, 1].value = i


The context manager can be replicated manually with

::

    try:

        xloil.app().ScreenUpdating = False
        xloil.app().EnableEvents = False
        xloil.app().Calculation = xloil.constants.xlCalculationManual

        ...

    finally:
    
        xloil.app().ScreenUpdating = True
        xloil.app().EnableEvents = True
        xloil.app().Calculation = xloil.constants.xlCalculationAutomatic


Troubleshooting
---------------

Both *comtypes* and *win32com* have caches for the python code backing the Excel object model. If 
these caches somehow become corrupted, it can result in strange COM errors.  It is safe to delete 
these caches and let the library regenerate them. The caches are located at:

   * *comtypes*: `<your python install>/site-packages/comtypes/gen`
   * *win32com*: run ``import win32com; print(win32com.__gen_path__)``

See `for example <https://stackoverflow.com/questions/52889704/>`_

Note: as of 25-Jan-2022, *comtypes* has been observed to give the wrong answer for a call to
`xloil.app().Workbooks(...)` so it is no longer used as the default whilst this is investigated.

"""
        The Python plugin for xlOil primarily allows creation of Excel functions and macros 
        backed by Python code. In addition it offers full control over GUI objects and an 
        interface for Excel automation: driving the application in code.

        See the documentation at https://xloil.readthedocs.io
      """
from __future__ import annotations
import typing

__all__ = [
    "Application",
    "Caller",
    "CannotConvert",
    "CellError",
    "ComBusyError",
    "ExcelArray",
    "ExcelGUI",
    "ExcelState",
    "ExcelWindow",
    "ExcelWindows",
    "ExcelWindowsIter",
    "IPyFromExcel",
    "IPyToExcel",
    "ObjectCache",
    "Range",
    "RangeIter",
    "RibbonControl",
    "RtdPublisher",
    "RtdReturn",
    "RtdServer",
    "StatusBar",
    "TaskPaneFrame",
    "Workbook",
    "Workbooks",
    "WorkbooksIter",
    "Worksheet",
    "Worksheets",
    "WorksheetsIter",
    "active_workbook",
    "active_worksheet",
    "app",
    "cache",
    "call",
    "call_async",
    "create_gui",
    "deregister_functions",
    "event",
    "excel_callback",
    "excel_state",
    "from_excel_date",
    "get_async_loop",
    "in_wizard",
    "insert_cell_image",
    "register_functions",
    "run",
    "run_async"
]


class Application():
    """
    Manages a handle to the *Excel.Application* object. This object is the root 
    of Excel's COM interface and supports a wide range of operations.

    In addition to the methods known to python, properties and methods of the 
    Application object can be resolved dynamically at runtime. The available methods
    will be familiar to VBA programmers and are well documented by Microsoft, 
    see `Object Model Overview <https://docs.microsoft.com/en-us/visualstudio/vsto/excel-object-model-overview>`_

    Note COM methods and properties are in UpperCamelCase, whereas python ones are 
    lower_case.

    Examples
    --------

    To get the name of the active worksheet:

        return xlo.app().ActiveWorksheet.Name

    See `Excel.Application <https://docs.microsoft.com/en-us/office/vba/api/excel.Application(object)>`_ 
    """
    def __enter__(self) -> object: ...
    def __exit__(self, arg0: object, arg1: object, arg2: object) -> None: ...
    def __getattr__(self, arg0: str) -> object: ...
    def __init__(self, com: object = None, hwnd: object = None, workbook: object = None) -> None: 
        """
        Creates a new Excel Application if no arguments are specified. Gets a handle to 
        an existing COM Application object based on the arguments.

        To get the parent Excel application if xlOil is embedded, used `xloil.app()`.

        Parameters
        ----------

        com: 
          Gets a handle to the given com object with class Excel.Appliction (marshalled 
          by `comtypes` or `win32com`).
        hwnd:
          Tries to gets a handle to the Excel.Application with given main window handle.
        workbook:
          Tries to gets a handle to the Excel.Application which has the specified workbook
          open.
        """
    def calculate(self, full: bool = False, rebuild: bool = False) -> None: 
        """
        Calculates all open workbooks

        Parameters
        ----------
        full:
          Forces a full calculation of the data in all open workbooks
        rebuild:
          For all open workbooks, forces a full calculation of the data 
          and rebuilds the dependencies. (Implies `full`)
        """
    def open(self, filepath: str, update_links: bool = True, read_only: bool = False) -> Workbook: 
        """
        Opens a workbook given its full `filepath`.

        Parameters
        ----------

        filepath: 
          path and filename of the target workbook
        update_links: 
          if True, attempts to update links to external workbooks
        read_only: 
          if True, opens the workbook in read-only mode
        """
    def quit(self, silent: bool = True) -> None: 
        """
        Terminates the application. If `silent` is True, unsaved data
        in workbooks is discarded, otherwise a prompt is displayed.
        """
    def range(self, address: str) -> object: 
        """
        Create a range object from an external address such as "[Book]Sheet!A1"
        """
    def to_com(self, lib: str = '') -> object: 
        """
        Returns a managed COM object which can be used to invoke Excel's full 
        object model. For details of the available calls see the Microsoft 
        documentation on the Excel Object Model. The ``lib`` used to provide COM
        support can be 'comtypes' or 'win32com'. If omitted, the default is 
        'win32com', unless specified in the XLL's ini file.
        """
    @property
    def enable_events(self) -> bool:
        """
                    Pauses or resumes Excel's event handling. It can be useful when writing to a sheet
                    to pause events both for performance and to prevent side effects.
                  

        :type: bool
        """
    @enable_events.setter
    def enable_events(self, arg1: bool) -> None:
        """
        Pauses or resumes Excel's event handling. It can be useful when writing to a sheet
        to pause events both for performance and to prevent side effects.
        """
    @property
    def visible(self) -> bool:
        """
                    Determines whether the Excel window is visble on the desktop
                  

        :type: bool
        """
    @visible.setter
    def visible(self, arg1: bool) -> None:
        """
        Determines whether the Excel window is visble on the desktop
        """
    @property
    def windows(self) -> ExcelWindows:
        """
        A collection of all Windows open in this Application

        :type: ExcelWindows
        """
    @property
    def workbooks(self) -> Workbooks:
        """
        A collection of all Workbooks open in this Application

        :type: Workbooks
        """
    pass
class Caller():
    """
    Captures the caller information for a worksheet function. On construction
    the class queries Excel via the `xlfCaller` function to determine the 
    calling cell or range. If the function was not called from a sheet (e.g. 
    via a macro), most of the methods return `None`.
    """
    def __init__(self) -> None: ...
    def __str__(self, arg0: bool) -> str: ...
    def address(self, a1style: bool = False) -> str: 
        """
        Gives the sheet address either in A1 form: '[Book]Sheet!A1' or RC form: '[Book]Sheet!R1C1'
        """
    @property
    def range(self) -> object:
        """
        Range object corresponding to caller address

        :type: object
        """
    @property
    def sheet_name(self) -> object:
        """
        Gives the sheet name of the caller or None if not called from a sheet.

        :type: object
        """
    @property
    def workbook(self) -> object:
        """
                    Gives the workbook name of the caller or None if not called from a sheet.
                    If the workbook has been saved, the name will contain a file extension.
                  

        :type: object
        """
    pass
class CannotConvert(Exception, BaseException):
    """
    Should be thrown by a converter when it is unable to handle the 
    provided type.  In a return converter it may not indicate a fatal 
    condition, as xlOil will fallback to another converter.
    """
    pass
class CellError():
    """
                  Enum-type class which represents an Excel error condition of the 
                  form `#N/A!`, `#NAME!`, etc passed as a function argument. If a 
                  function argument does not specify a type (e.g. int, str) it may be passed 
                  a CellError, which it can handle based on the error condition.
                

    Members:

      NULL

      DIV

      VALUE

      REF

      NAME

      NUM

      NA

      GETTING_DATA
    """
    def __eq__(self, other: object) -> bool: ...
    def __getstate__(self) -> int: ...
    def __hash__(self) -> int: ...
    def __index__(self) -> int: ...
    def __init__(self, value: int) -> None: ...
    def __int__(self) -> int: ...
    def __ne__(self, other: object) -> bool: ...
    def __repr__(self) -> str: ...
    def __setstate__(self, state: int) -> None: ...
    @property
    def name(self) -> str:
        """
        :type: str
        """
    @property
    def value(self) -> int:
        """
        :type: int
        """
    DIV: xloil_core.CellError = None # value = <CellError.DIV: 7>
    GETTING_DATA: xloil_core.CellError = None # value = <CellError.GETTING_DATA: 43>
    NA: xloil_core.CellError = None # value = <CellError.NA: 42>
    NAME: xloil_core.CellError = None # value = <CellError.NAME: 29>
    NULL: xloil_core.CellError = None # value = <CellError.NULL: 0>
    NUM: xloil_core.CellError = None # value = <CellError.NUM: 36>
    REF: xloil_core.CellError = None # value = <CellError.REF: 23>
    VALUE: xloil_core.CellError = None # value = <CellError.VALUE: 15>
    __members__: dict = None # value = {'NULL': <CellError.NULL: 0>, 'DIV': <CellError.DIV: 7>, 'VALUE': <CellError.VALUE: 15>, 'REF': <CellError.REF: 23>, 'NAME': <CellError.NAME: 29>, 'NUM': <CellError.NUM: 36>, 'NA': <CellError.NA: 42>, 'GETTING_DATA': <CellError.GETTING_DATA: 43>}
    pass
class ComBusyError(Exception, BaseException):
    pass
class ExcelArray():
    """
    A view of a internal Excel array which can be manipulated without
    copying the underlying data. It's not a general purpose array class 
    but rather used to create efficiencies in type converters.

    It can be accessed and sliced using the usual syntax (the slice step must be 1):

    ::

        x[1, 1] # The value at 1,1 as int, str, float, etc.

        x[1, :] # The second row as another ExcelArray

        x[:-1, :-1] # A sub-array omitting the last row and column
    """
    def __getitem__(self, arg0: tuple) -> object: 
        """
        Given a 2-tuple, slices the array to return a sub ExcelArray or a single element.
        """
    def slice(self, from_row: int, from_col: int, to_row: int, to_col: int) -> ExcelArray: 
        """
        Slices the array 
        """
    def to_numpy(self, dtype: typing.Optional[int] = None, dims: typing.Optional[int] = 2) -> object: 
        """
        Converts the array to a numpy array. If *dtype* is None, xlOil attempts 
        to determine the correct numpy dtype. It raises an exception if values
        cannot be converted to a specified *dtype*. The array dimension *dims* 
        can be 1 or 2 (default is 2).
        """
    @property
    def dims(self) -> int:
        """
        Property which gives the dimension of the array: 1 or 2

        :type: int
        """
    @property
    def ncols(self) -> int:
        """
        Returns the number of columns in the array

        :type: int
        """
    @property
    def nrows(self) -> int:
        """
        Returns the number of rows in the array

        :type: int
        """
    @property
    def shape(self) -> tuple:
        """
        Returns a tuple (nrows, ncols) like numpy's array.shape

        :type: tuple
        """
    pass
class ExcelGUI():
    """
    Controls an Ribbon and its associated COM addin. The methods of this object are safe
    to call from any thread.  However, COM must be used on Excel's main thread, so the methods  
    schedule calls to run on the main thead. This could lead to deadlocks if the call 
    triggers event handlers on the main thread, which in turn block whilst waiting for the 
    thread originally calling ExcelGUI.
    """
    def activate(self, id: str) -> bool: 
        """
        Activatives the ribbon tab with the specified id.  Returns False if
        there is no Ribbon or the Ribbon is collapsed.
        """
    def connect(self, xml: str = '', func_names: object = None) -> _Future: 
        """
        Connects this COM add-in underlying this Ribbon to Excel. Any specified 
        ribbon XML will be passed to Excel.
        """
    def create_task_pane(self, *args, **kwargs) -> object: 
        """
        Returns a task pane with title <name> attached to the active window,
        creating it if it does not already exist.  See `xloil.create_task_pane`.

        Parameters
        ----------

        creator: 
            * a subclass of `QWidget` or
            * a function which takes a `TaskPaneFrame` and returns a `CustomTaskPane`

        window: 
            a window title or `ExcelWindow` object to which the task pane should be
            attached.  If None, the active window is used.
        """
    def disconnect(self) -> None: 
        """
        Unloads the underlying COM add-in and any ribbon customisation.
        """
    def invalidate(self, id: str = '') -> None: 
        """
        Invalidates the specified control: this clears the cache of responses
        to callbacks associated with the control. For example, this can be
        used to hide a control by forcing its getVisible callback to be invoked,
        rather than using the cached value.

        If no control ID is specified, all controls are invalidated.
        """
    def task_pane_frame(self, name: str, progid: object = None, window: object = None) -> _CTPFuture: 
        """
        Used internally to create a custom task pane window which can be populated
        with a python GUI.  Most users should use `create_task_pane(...)` instead.

        A COM `progid` can be specified, but this will prevent using a python GUI
        in the task pane. This is a specialised use case.
        """
    @property
    def name(self) -> str:
        """
        :type: str
        """
    pass
class ExcelState():
    """
    Gives information about the Excel application. Cannot be constructed: call
    ``xloil.excel_state`` to get an instance.
    """
    @property
    def hinstance(self) -> capsule:
        """
        Excel Win32 HINSTANCE

        :type: capsule
        """
    @property
    def hwnd(self) -> int:
        """
        Excel Win32 main window handle(as an int)

        :type: int
        """
    @property
    def main_thread_id(self) -> int:
        """
        Excel main thread ID

        :type: int
        """
    @property
    def version(self) -> int:
        """
        Excel major version

        :type: int
        """
    pass
class ExcelWindow():
    """
    Represents a window.  A window is a view of a workbook.
    See `Excel.Window <https://docs.microsoft.com/en-us/office/vba/api/excel.WindowWindow>`_ 
    """
    def __getattr__(self, arg0: str) -> object: ...
    def __str__(self) -> str: ...
    def to_com(self, lib: str = '') -> object: 
        """
        Returns a managed COM object which can be used to invoke Excel's full 
        object model. For details of the available calls see the Microsoft 
        documentation on the Excel Object Model. The ``lib`` used to provide COM
        support can be 'comtypes' or 'win32com'. If omitted, the default is 
        'win32com', unless specified in the XLL's ini file.
        """
    @property
    def app(self) -> Application:
        """
                  Returns the parent `xloil.Application` object associated with this object.
              

        :type: Application
        """
    @property
    def hwnd(self) -> int:
        """
        The Win32 API window handle as an integer

        :type: int
        """
    @property
    def name(self) -> str:
        """
        :type: str
        """
    @property
    def workbook(self) -> Workbook:
        """
        The workbook being displayed by this window

        :type: Workbook
        """
    pass
class ExcelWindows():
    """
    A collection of all the Window objects in Excel.  A Window is a view of
    a Workbook

    See `Excel.Windows <https://docs.microsoft.com/en-us/office/vba/api/excel.WindowsWindows>`_ 
    """
    def __getitem__(self, arg0: str) -> ExcelWindow: ...
    def __iter__(self) -> ExcelWindowsIter: ...
    def __len__(self) -> int: ...
    def get(self, name: str, default: object = None) -> object: 
        """
        Tries to get the named object, returning the default if not found
        """
    @property
    def active(self) -> object:
        """
                      Gives the active (as displayed in the GUI) object in the collection
                      or None if no object has been activated.
                    

        :type: object
        """
    pass
class ExcelWindowsIter():
    def __iter__(self) -> object: ...
    def __next__(self) -> ExcelWindow: ...
    pass
class IPyFromExcel():
    def __call__(self, arg0: object) -> None: ...
    pass
class IPyToExcel():
    pass
class ObjectCache():
    """
    Provides a way to manipulate xlOil's Python object cache

    Examples
    --------

    ::

        @xlo.func
        def myfunc(x):
            return xlo.cache(MyObject(x)) # <-equivalent to cache.add(...)

        @xlo.func
        def myfunc2(array: xlo.Array(str), i):
            return xlo.cache[array[i]]   # <-equivalent to cache.get(...)
    """
    def __call__(self, obj: object, tag: str = '', key: str = '') -> object: 
        """
        Calls `add` method with provided arguments
        """
    def __contains__(self, arg0: str) -> bool: ...
    def __getitem__(self, arg0: str) -> object: ...
    def add(self, obj: object, tag: str = '', key: str = '') -> object: 
        """
        Adds an object to the cache and returns a reference string
        based on the currently calculating cell.

        xlOil automatically adds unconvertible returned objects to the cache,
        so this function is useful to force a recognised object, such as an 
        iterable into the cache, or to return a list of cached objects.

        Parameters
        ----------

        obj:
          The object to cache.  Required.

        tag: str
          An optional string to append to the cache ref to make it more 
          'friendly'. When returning python objects from functions, 
          xlOil uses the object's type name as a tag

        key: str
          If specified, use the exact cache key (after prepending by
          cache uniquifier). The user is responsible for ensuring 
          uniqueness of the cache key.
        """
    def contains(self, ref: str) -> bool: 
        """
        Returns True if the given reference string links to a valid object
        """
    def get(self, ref: str, default: object = None) -> object: 
        """
        Fetches an object from the cache given a reference string.
        Returns `default` if not found
        """
    def keys(self) -> list: 
        """
        Returns all cache keys as a list of strings
        """
    def remove(self, ref: str) -> bool: ...
    pass
class Range():
    """
    Represents a cell, a row, a column or a selection of cells containing a contiguous 
    blocks of cells. (Non contiguous ranges are not currently supported).
    This class allows direct access to an area on a worksheet. It uses similar 
    syntax to Excel's Range object, supporting the ``cell`` and ``range`` functions,  
    however indices are zero-based as per python's standard.

    A Range can be accessed and sliced using the usual syntax (the slice step must be 1):

    ::

        x[1, 1] # The value at (1, 1) as a python type: int, str, float, etc.

        x[1, :] # The second row as another Range object

        x[:-1, :-1] # A sub-range omitting the last row and column

    See `Excel.Range <https://docs.microsoft.com/en-us/office/vba/api/excel.Range(object)>`_ 
    """
    def __getattr__(self, arg0: str) -> object: ...
    def __getitem__(self, arg0: object) -> object: 
        """
        Given a 2-tuple, slices the range to return a sub Range or a single element.Uses
        normal python slicing conventions i.e[left included, right excluded), negative
        numbers are offset from the end.If the tuple specifies a single cell, returns
        the value in that cell, otherwise returns a Range object.
        """
    def __init__(self, address: str) -> None: ...
    def __iter__(self) -> RangeIter: ...
    def __len__(self) -> int: ...
    def __str__(self) -> str: ...
    def address(self, local: bool = False) -> str: 
        """
        Returns the address of the range in A1 format, e.g. *[Book]SheetNm!A1:Z5*. The 
        sheet name may be surrounded by single quote characters if it contains a space.
        If *local* is set to true, everything prior to the '!' is omitted.
        """
    def cell(self, row: int, col: int) -> Range: 
        """
        Returns a Range object which consists of a single cell. The indices are zero-based 
        from the top left of the parent range.
        """
    def clear(self) -> None: 
        """
        Clears all values and formatting.  Any cell in the range will then have Empty type.
        """
    def range(self, from_row: int, from_col: int, to_row: object = None, to_col: object = None, num_rows: object = None, num_cols: object = None) -> Range: 
        """
        Creates a subrange using offsets from the top left corner of the parent range.
        Like Excel's Range function, we allow negative offsets to select ranges outside the
        parent.

        Parameters
        ----------

        from_row: int
            Starting row offset from the top left of the parent range. Zero-based. Can be negative

        from_col: int
            Starting row offset from the top left of the parent range. Zero-based. Can be negative

        to_row: int
            End row offset from the top left of the parent range. This row will be *included* in 
            the range. The offset is zero-based and can be negative to select ranges outside the
            parent range. Do not specify both `to_row` and `num_rows`.

        to_col: int
            End column offset from the top left of the parent range. This column will be *included*
            in the range. The offset is zero-based and can be negative to select ranges outside 
            the parent range. Do not specify both `to_col` and `num_cols`.

        num_rows: int
            Number of rows in output range. Must be positive. If neither `num_rows` or `to_rows` 
            are specified, the range ends at the last row of the parent range.

        num_cols: int
            Number of columns in output range. Must be positive. If neither `num_cols` or `to_cols` 
            are specified, the range ends at the last column of the parent range.
        """
    def set(self, arg0: object) -> None: 
        """
        Sets the data in the range to the provided value. If a single value is passed
        all cells will be set to the value. If a 2d-array is provided, the array will be
        pasted at the top-left of the range with the remainging cells being set to #N/A.
        If a 1d array is provided it will be pasted at the top left and repeated down or
        right depending on orientation.
        """
    def to_com(self, lib: str = '') -> object: 
        """
        Returns a managed COM object which can be used to invoke Excel's full 
        object model. For details of the available calls see the Microsoft 
        documentation on the Excel Object Model. The ``lib`` used to provide COM
        support can be 'comtypes' or 'win32com'. If omitted, the default is 
        'win32com', unless specified in the XLL's ini file.
        """
    def trim(self) -> Range: 
        """
        Returns a sub-range by trimming to the last non-empty (i.e. not Nil, #N/A or "") 
        row and column. The top-left remains the same so the function always returns
        at least a single cell, even if it's empty.  
        """
    @property
    def bounds(self) -> typing.Tuple[int, int, int, int]:
        """
                    Returns a zero-based tuple (top-left-row, top-left-col, bottom-right-row, bottom-right-col)
                    which defines the Range area (currently only rectangular ranges are supported).
                  

        :type: typing.Tuple[int, int, int, int]
        """
    @property
    def formula(self) -> str:
        """
                    Get / sets the forumula for the range as a string string. If the range
                    is larger than one cell, the formula is applied as an ArrayFormula.
                    Returns an empty string if the range does not contain a formula or array 
                    formula.
                  

        :type: str
        """
    @formula.setter
    def formula(self, arg1: str) -> None:
        """
        Get / sets the forumula for the range as a string string. If the range
        is larger than one cell, the formula is applied as an ArrayFormula.
        Returns an empty string if the range does not contain a formula or array 
        formula.
        """
    @property
    def ncols(self) -> int:
        """
        Returns the number of columns in the range

        :type: int
        """
    @property
    def nrows(self) -> int:
        """
        Returns the number of rows in the range

        :type: int
        """
    @property
    def parent(self) -> Worksheet:
        """
        Returns the parent Worksheet for this Range

        :type: Worksheet
        """
    @property
    def shape(self) -> typing.Tuple[int, int]:
        """
        Returns a tuple (num columns, num rows)

        :type: typing.Tuple[int, int]
        """
    @property
    def value(self) -> object:
        """
                    Property which gets or sets the value for a range. A fetched value is converted
                    to the most appropriate Python type using the normal generic converter.

                    If you use a horizontal array for the assignment, it is duplicated down to fill 
                    the entire rectangle. If you use a vertical array, it is duplicated right to fill 
                    the entire rectangle. If you use a rectangular array, and it is too small for the 
                    rectangular range you want to put it in, that range is padded with #N/As.
                  

        :type: object
        """
    @value.setter
    def value(self, arg1: object) -> None:
        """
        Property which gets or sets the value for a range. A fetched value is converted
        to the most appropriate Python type using the normal generic converter.

        If you use a horizontal array for the assignment, it is duplicated down to fill 
        the entire rectangle. If you use a vertical array, it is duplicated right to fill 
        the entire rectangle. If you use a rectangular array, and it is too small for the 
        rectangular range you want to put it in, that range is padded with #N/As.
        """
    pass
class RangeIter():
    def __iter__(self) -> object: ...
    def __next__(self) -> _object: ...
    pass
class RibbonControl():
    """
    This object is passed to ribbon callback handlers to indicate which control  
    raised the callback.
    """
    @property
    def id(self) -> str:
        """
        A string that represents the Id attribute for the control or custom menu item

        :type: str
        """
    @property
    def tag(self) -> str:
        """
        A string that represents the Tag attribute for the control or custom menu item.

        :type: str
        """
    pass
class RtdPublisher():
    """
    RTD servers use a publisher/subscriber model with the 'topic' as the key
    The publisher class is linked to a single topic string.

    Typically the publisher will do nothing on construction, but when it detects
    a subscriber using the connect() method, it creates a background publishing task
    When disconnect() indicates there are no subscribers, it cancels this task with
    a call to stop()

    If the task is slow to return or spin up, it could be started the constructor  
    and kept it running permanently, regardless of subscribers.

    The publisher should call RtdServer.publish() to push values to subscribers.
    """
    def __init__(self) -> None: 
        """
        This __init__ method must be called explicitly by subclasses or pybind
        will fatally crash Excel.
        """
    def connect(self, num_subscribers: int) -> None: 
        """
        Called by the RtdServer when a sheet function subscribes to this 
        topic. Typically a topic will start up its publisher on the first
        subscriber, i.e. when num_subscribers == 1
        """
    def disconnect(self, num_subscribers: int) -> bool: 
        """
        Called by the RtdServer when a sheet function disconnects from this 
        topic. This happens when the function arguments are changed the
        function deleted. Typically a topic will shutdown its publisher 
        when num_subscribers == 0.

        Whilst the topic remains live, it may still receive new connection
        requests, so generally avoid finalising in this method.
        """
    def done(self) -> bool: 
        """
        Returns True if the topic can safely be deleted without leaking resources.
        """
    def stop(self) -> None: 
        """
        Called by the RtdServer to indicate that a topic should shutdown
        and dependent threads or tasks and finalise resource usage
        """
    @property
    def topic(self) -> str:
        """
        Returns the name of the topic

        :type: str
        """
    pass
class RtdReturn():
    def set_done(self) -> None: ...
    def set_result(self, arg0: object) -> None: ...
    def set_task(self, arg0: object) -> None: ...
    @property
    def caller(self) -> Caller:
        """
        :type: Caller
        """
    @property
    def loop(self) -> object:
        """
        :type: object
        """
    pass
class RtdServer():
    """
    An RtdServer allows asynchronous interaction with Excel by posting update
    notifications which cause Excel to recalcate certain cells.  The python 
    RtdServer object manages an RTD COM server with each new RtdServer creating
    an underlying COM server. The RtdServer works on a publisher/subscriber
    model with topics identified by a string. 

    A topic publisher is registered using start(). Subsequent calls to subscribe()
    will connect this topic and tell Excel that the current calling cell should be
    recalculated when a new value is published.

    RTD sits outside of Excel's normal calc cycle: publishers can publish new values 
    at any time, triggering a re-calc of any cells containing subscribers. Note the
    re-calc will only happen 'live' if Excel's caclulation mode is set to automatic
    """
    def __init__(self) -> None: ...
    def drop(self, arg0: str) -> None: 
        """
        Drops the producer for a topic by calling `RtdPublisher.stop()`, then waits
        for it to complete and publishes #N/A to all subscribers.
        """
    def peek(self, topic: str, converter: IPyFromExcel = None) -> object: 
        """
        Looks up a value for a specified topic, but does not subscribe.
        If there is no active publisher for the topic, it returns None.
        If there is no published value, it will return CellError.NA.

        This function does not use any Excel API and is safe to call at
        any time on any thread.
        """
    def publish(self, topic: str, value: object, converter: IPyToExcel = None) -> bool: 
        """
        Publishes a new value for the specified topic and updates all subscribers.
        This function can be called even if no RtdPublisher has been started.

        This function does not use any Excel API and is safe to call at any time
        on any thread.

        An Exception object can be passed at the value, this will trigger the debugging
        hook if it is set. The exception string and it's traceback will be published.
        """
    def start(self, topic: RtdPublisher) -> None: 
        """
        Registers an RtdPublisher publisher with this manager. The RtdPublisher receives
        notification when the number of subscribers changes
        """
    def start_task(self, topic: str, func: object, converter: IPyToExcel = None) -> None: 
        """
        Launch a publishing task for a `topic` given a func and a return converter
        """
    def subscribe(self, topic: str, converter: IPyFromExcel = None) -> object: 
        """
        Subscribes to the specified topic. If no publisher for the topic currently 
        exists, it returns None, but the subscription is held open and will connect
        to a publisher created later. If there is no published value, it will return 
        CellError.NA.  

        This calls Excel's RTD function, which means the calling cell will be
        recalculated every time a new value is published.

        Calling this function outside of a worksheet function called by Excel may
        produce undesired results and possibly crash Excel.
        """
    pass
class StatusBar():
    """
     
    Displays status bar messages and clears the status bar (after an optional delay) 
    on context exit.

    Examples
    --------

    with StatusBar(1000) as status:
      status.msg('Doing slow thing')
      ...
      status.msg('Done slow thing')
    """
    def __enter__(self) -> object: ...
    def __exit__(self, *args) -> None: ...
    def __init__(self, timeout: int = 0) -> None: 
        """
        Constructs a StatusBar with a timeout specified in milliseconds.  After the 
        StatusBar context exits, any messages will be cleared after the timeout
        """
    def msg(self, msg: str, timeout: int = 0) -> None: 
        """
        Posts a status bar message, and if `timeout` is non-zero, clears if after
        the specified number of milliseconds
        """
    pass
class TaskPaneFrame():
    """
    References Excel's base task pane object into which the python GUI can be drawn.
    The methods of this object are safe to call from any thread.  COM must be used on Excel's
    main thread, so the methods all wrap their calls to ensure to this happens. This could lead 
    to deadlocks if the call triggers event  handlers on the main thread, which in turn block 
    waiting for the thread originally calling `TaskPaneFrame`.
    """
    def add_event_handler(self, handler: object) -> None: ...
    def com_control(self, lib: str = '') -> object: 
        """
        Gets the base COM control of the task pane. The ``lib`` used to provide
        COM support can be 'comtypes' or 'win32com' (default is win32com).
        """
    @property
    def parent_hwnd(self) -> int:
        """
        Win32 window handle used to attach a python GUI to a task pane frame

        :type: int
        """
    @property
    def size(self) -> typing.Tuple[int, int]:
        """
        Gets/sets the task pane size as a tuple (width, height)

        :type: typing.Tuple[int, int]
        """
    @size.setter
    def size(self, arg1: tuple) -> None:
        """
        Gets/sets the task pane size as a tuple (width, height)
        """
    @property
    def title(self) -> str:
        """
        :type: str
        """
    @property
    def visible(self) -> bool:
        """
        Determines the visibility of the task pane

        :type: bool
        """
    @visible.setter
    def visible(self, arg1: bool) -> None:
        """
        Determines the visibility of the task pane
        """
    @property
    def window(self) -> ExcelWindow:
        """
        Gives the window of the document window to which the frame is attached, can be used to uniquely identify the pane

        :type: ExcelWindow
        """
    pass
class Workbook():
    """
    Represents an open Excel workbook.
    See `Excel.Workbook <https://docs.microsoft.com/en-us/office/vba/api/excel.WorkbookWorkbook>`_ 
    """
    def __enter__(self) -> object: ...
    def __exit__(self, arg0: object, arg1: object, arg2: object) -> None: ...
    def __getattr__(self, arg0: str) -> object: ...
    def __getitem__(self, arg0: object) -> object: 
        """
        If the index is a worksheet name, returns the `Worksheet` object,
        otherwise treats the string as a workbook address and returns a `Range`.
        """
    def __str__(self) -> str: ...
    def add(self, name: object = None, before: object = None, after: object = None) -> Worksheet: 
        """
        Add a worksheet, returning a `Worksheet` object.

        Parameters
        ----------
        name: str
          Names the worksheet, otherwise it will have an Excel-assigned name
        before: Worksheet
          Places the new worksheet immediately before this Worksheet object 
        after: Worksheet
          Places the new worksheet immediately before this Worksheet object.
          Specifying both `before` and `after` raises an exception.
        """
    def close(self, save: bool = True) -> None: 
        """
        Closes the workbook. If there are changes to the workbook and the 
        workbook doesn't appear in any other open windows, the `save` argument
        specifies whether changes should be saved. If set to True, changes are 
        saved to the workbook, if False they are discared.
        """
    def range(self, address: str) -> object: 
        """
        Create a `Range` object from an address such as "Sheet!A1" or a named range
        """
    def save(self, filepath: str = '') -> None: 
        """
        Saves the Workbook, either to the specified `filepath` or if this is
        unspecified, to its original source file (an error is raised if the 
        workbook has never been saved).
        """
    def to_com(self, lib: str = '') -> object: 
        """
        Returns a managed COM object which can be used to invoke Excel's full 
        object model. For details of the available calls see the Microsoft 
        documentation on the Excel Object Model. The ``lib`` used to provide COM
        support can be 'comtypes' or 'win32com'. If omitted, the default is 
        'win32com', unless specified in the XLL's ini file.
        """
    def worksheet(self, name: str) -> Worksheet: 
        """
        Returns the named worksheet which is part of this workbook (if it exists)
        otherwise raises an exception.
        """
    @property
    def app(self) -> Application:
        """
                  Returns the parent `xloil.Application` object associated with this object.
              

        :type: Application
        """
    @property
    def name(self) -> str:
        """
        :type: str
        """
    @property
    def path(self) -> str:
        """
        The full path to the workbook, including the filename

        :type: str
        """
    @property
    def windows(self) -> ExcelWindows:
        """
                    A collection object of all windows which are displaying this workbook
                  

        :type: ExcelWindows
        """
    @property
    def worksheets(self) -> Worksheets:
        """
                    A collection object of all worksheets which are part of this workbook
                  

        :type: Worksheets
        """
    pass
class Workbooks():
    """
    A collection of all the Workbook objects that are currently open in the 
    Excel application.  

    See `Excel.Workbooks <https://docs.microsoft.com/en-us/office/vba/api/excel.WorkbooksWorkbooks>`_ 
    """
    def __getitem__(self, arg0: str) -> Workbook: ...
    def __iter__(self) -> WorkbooksIter: ...
    def __len__(self) -> int: ...
    def add(self) -> Workbook: 
        """
        Creates and returns a new workbook with an Excel-assigned name
        """
    def get(self, name: str, default: object = None) -> object: 
        """
        Tries to get the named object, returning the default if not found
        """
    @property
    def active(self) -> object:
        """
                      Gives the active (as displayed in the GUI) object in the collection
                      or None if no object has been activated.
                    

        :type: object
        """
    pass
class WorkbooksIter():
    def __iter__(self) -> object: ...
    def __next__(self) -> Workbook: ...
    pass
class Worksheet():
    """
    Allows access to ranges and properties of a worksheet. It uses similar 
    syntax to Excel's Worksheet object, supporting the ``cell`` and ``range`` functions, 
    however indices are zero-based as per python's standard.

    See `Excel.Worksheet <https://docs.microsoft.com/en-us/office/vba/api/excel.WorksheetWorksheet>`_ 
    """
    def __getattr__(self, arg0: str) -> object: ...
    def __getitem__(self, arg0: object) -> object: 
        """
        If the argument is a string, returns the range specified by the local address, 
        equivalent to ``at_address``.  

        If the argument is a 2-tuple, slices the sheet to return a Range or a single element. 
        Uses normal python slicing conventions, i.e [left included, right excluded), negative
        numbers are offset from the end.
        """
    def __str__(self) -> str: ...
    def activate(self) -> None: 
        """
        Makes this worksheet the active sheet
        """
    def at(self, address: str) -> Range: 
        """
        Returns the range specified by the local address, e.g. ``.at('B3:D6')``
        """
    def calculate(self) -> None: 
        """
        Calculates this worksheet
        """
    def cell(self, row: int, col: int) -> Range: 
        """
        Returns a Range object which consists of a single cell. The indices are zero-based 
        from the top left of the parent range.
        """
    def range(self, from_row: int, from_col: int, to_row: object = None, to_col: object = None, num_rows: object = None, num_cols: object = None) -> Range: 
        """
        Specifies a range in this worksheet.

        Parameters
        ----------

        from_row: int
            Starting row offset from the top left of the parent range. Zero-based.

        from_col: int
            Starting row offset from the top left of the parent range. Zero-based.

        to_row: int
            End row offset from the top left of the parent range. This row will be *included* in 
            the range. The offset is zero-based. Do not specify both `to_row` and `num_rows`.

        to_col: int
            End column offset from the top left of the parent range. This column will be *included*  
            in the range. The offset is zero-based. Do not specify both `to_col` and `num_cols`.

        num_rows: int
            Number of rows in output range. Must be positive. If neither `num_rows` or `to_rows` 
            are specified, the range ends at the end of the sheet.

        num_cols: int
            Number of columns in output range. Must be positive. If neither `num_cols` or `to_cols` 
            are specified, the range ends at the end of the sheet.
        """
    def to_com(self, lib: str = '') -> object: 
        """
        Returns a managed COM object which can be used to invoke Excel's full 
        object model. For details of the available calls see the Microsoft 
        documentation on the Excel Object Model. The ``lib`` used to provide COM
        support can be 'comtypes' or 'win32com'. If omitted, the default is 
        'win32com', unless specified in the XLL's ini file.
        """
    @property
    def app(self) -> Application:
        """
                  Returns the parent `xloil.Application` object associated with this object.
              

        :type: Application
        """
    @property
    def name(self) -> str:
        """
        :type: str
        """
    @property
    def parent(self) -> Workbook:
        """
        Returns the parent Workbook for this Worksheet

        :type: Workbook
        """
    pass
class Worksheets():
    """
    A collection of all the Worksheet objects in the specified or active workbook. 
    Each Worksheet object represents a worksheet.

    See `Excel.Worksheets <https://docs.microsoft.com/en-us/office/vba/api/excel.WorksheetsWorksheets>`_ 
    """
    def __getitem__(self, arg0: str) -> Worksheet: ...
    def __iter__(self) -> WorksheetsIter: ...
    def __len__(self) -> int: ...
    def add(self, name: object = None, before: object = None, after: object = None) -> Worksheet: 
        """
        Add a worksheet, returning a `Worksheet` object.

        Parameters
        ----------
        name: str
          Names the worksheet, otherwise it will have an Excel-assigned name
        before: Worksheet
          Places the new worksheet immediately before this Worksheet object 
        after: Worksheet
          Places the new worksheet immediately before this Worksheet object.
          Specifying both `before` and `after` raises an exception.
        """
    def get(self, name: str, default: object = None) -> object: 
        """
        Tries to get the named object, returning the default if not found
        """
    @property
    def active(self) -> object:
        """
                      Gives the active (as displayed in the GUI) object in the collection
                      or None if no object has been activated.
                    

        :type: object
        """
    pass
class WorksheetsIter():
    def __iter__(self) -> object: ...
    def __next__(self) -> Worksheet: ...
    pass
class _AddinFuture():
    """
    A Future represents an eventual result of an asynchronous operation.
    Future is an awaitable object. Coroutines can await on Future objects 
    until they either have a result or an exception set. This Future cannot
    be cancelled.

    This class actually wraps a C++ future so does executes in a separate 
    thread unrelated to an `asyncio` event loop. 
    """
    def __await__(self) -> _AddinFutureIter: ...
    def done(self) -> bool: 
        """
        Return True if the Future is done.  A Future is done if it has a result or an exception.
        """
    def result(self) -> object: 
        """
        Return the result of the Future, blocking if the Future is not yet done.

        If the Future has a result, its value is returned.

        If the Future has an exception, raises the exception.
        """
    pass
class _AddinFutureIter():
    def __iter__(self) -> object: ...
    def __next__(self) -> None: ...
    pass
class _AsyncReturn():
    def set_done(self) -> None: ...
    def set_result(self, arg0: object) -> None: ...
    def set_task(self, arg0: object) -> None: ...
    @property
    def caller(self) -> Caller:
        """
        :type: Caller
        """
    @property
    def loop(self) -> object:
        """
        :type: object
        """
    pass
class _CTPFuture():
    """
    A Future represents an eventual result of an asynchronous operation.
    Future is an awaitable object. Coroutines can await on Future objects 
    until they either have a result or an exception set. This Future cannot
    be cancelled.

    This class actually wraps a C++ future so does executes in a separate 
    thread unrelated to an `asyncio` event loop. 
    """
    def __await__(self) -> _CTPFutureIter: ...
    def done(self) -> bool: 
        """
        Return True if the Future is done.  A Future is done if it has a result or an exception.
        """
    def result(self) -> object: 
        """
        Return the result of the Future, blocking if the Future is not yet done.

        If the Future has a result, its value is returned.

        If the Future has an exception, raises the exception.
        """
    pass
class _CTPFutureIter():
    def __iter__(self) -> object: ...
    def __next__(self) -> None: ...
    pass
class _CustomConverter(IPyFromExcel):
    """
    This is the interface class for custom type converters to allow them
    to be called from the Core.
    """
    def __init__(self, callable: object, check_cache: bool = True) -> None: ...
    pass
class _CustomReturn(IPyToExcel):
    def __init__(self, callable: object) -> None: ...
    def get_handler(self) -> object: ...
    pass
class _CustomReturnConverter():
    @property
    def value(self) -> IPyToExcel:
        """
        :type: IPyToExcel
        """
    @value.setter
    def value(self, arg0: IPyToExcel) -> None:
        pass
    pass
class _ExcelObjFuture():
    """
    A Future represents an eventual result of an asynchronous operation.
    Future is an awaitable object. Coroutines can await on Future objects 
    until they either have a result or an exception set. This Future cannot
    be cancelled.

    This class actually wraps a C++ future so does executes in a separate 
    thread unrelated to an `asyncio` event loop. 
    """
    def __await__(self) -> _ExcelObjFutureIter: ...
    def done(self) -> bool: 
        """
        Return True if the Future is done.  A Future is done if it has a result or an exception.
        """
    def result(self) -> object: 
        """
        Return the result of the Future, blocking if the Future is not yet done.

        If the Future has a result, its value is returned.

        If the Future has an exception, raises the exception.
        """
    pass
class _ExcelObjFutureIter():
    def __iter__(self) -> object: ...
    def __next__(self) -> None: ...
    pass
class _FuncArg():
    def __init__(self) -> None: ...
    @property
    def allow_range(self) -> bool:
        """
        :type: bool
        """
    @allow_range.setter
    def allow_range(self, arg0: bool) -> None:
        pass
    @property
    def converter(self) -> IPyFromExcel:
        """
        :type: IPyFromExcel
        """
    @converter.setter
    def converter(self, arg0: IPyFromExcel) -> None:
        pass
    @property
    def default(self) -> object:
        """
        :type: object
        """
    @default.setter
    def default(self, arg0: object) -> None:
        pass
    @property
    def help(self) -> str:
        """
        :type: str
        """
    @help.setter
    def help(self, arg0: str) -> None:
        pass
    @property
    def name(self) -> str:
        """
        :type: str
        """
    @name.setter
    def name(self, arg0: str) -> None:
        pass
    pass
class _FuncSpec():
    def __init__(self, func: function, args: typing.List[_FuncArg], name: str = '', features: str = None, help: str = '', category: str = '', local: bool = True, volatile: bool = False, has_kwargs: bool = False) -> None: ...
    def __str__(self) -> str: ...
    @property
    def args(self) -> typing.List[_FuncArg]:
        """
        :type: typing.List[_FuncArg]
        """
    @property
    def help(self) -> str:
        """
        :type: str
        """
    @property
    def name(self) -> str:
        """
        :type: str
        """
    @property
    def return_converter(self) -> IPyToExcel:
        """
        :type: IPyToExcel
        """
    @return_converter.setter
    def return_converter(self, arg1: IPyToExcel) -> None:
        pass
    pass
class _Future():
    def __await__(self) -> _FutureIter: ...
    def done(self) -> bool: ...
    def result(self) -> object: ...
    pass
class _FutureIter():
    def __iter__(self) -> object: ...
    def __next__(self) -> None: ...
    pass
class _LogWriter():
    """
    Writes a log message to xlOil's log.  The level parameter can be a level constant 
    from the `logging` module or one of the strings *error*, *warn*, *info*, *debug* or *trace*.

    Only messages with a level higher than the xlOil log level which is initially set
    to the value in the xlOil settings will be output to the log file. Trace output
    can only be seen with a debug build of xlOil.
    """
    def __call__(self, msg: str, level: object = 20) -> None: 
        """
        Writes a message to the log at the optionally specifed level. The default 
        level is 'info'.
        """
    def __init__(self) -> None: 
        """
        Do not construct this class - a singleton instance is created by xlOil.
        """
    @property
    def level(self) -> int:
        """
                      Returns or sets the current log level. The returned value will always be an 
                      integer corresponding to levels in the `logging` module.  The level can be
                      set to an integer or one of the strings *error*, *warn*, *info*, *debug* or *trace*.
                    

        :type: int
        """
    @level.setter
    def level(self, arg1: object) -> None:
        """
        Returns or sets the current log level. The returned value will always be an 
        integer corresponding to levels in the `logging` module.  The level can be
        set to an integer or one of the strings *error*, *warn*, *info*, *debug* or *trace*.
        """
    pass
class _PyObjectFuture():
    """
    A Future represents an eventual result of an asynchronous operation.
    Future is an awaitable object. Coroutines can await on Future objects 
    until they either have a result or an exception set. This Future cannot
    be cancelled.

    This class actually wraps a C++ future so does executes in a separate 
    thread unrelated to an `asyncio` event loop. 
    """
    def __await__(self) -> _PyObjectFutureIter: ...
    def done(self) -> bool: 
        """
        Return True if the Future is done.  A Future is done if it has a result or an exception.
        """
    def result(self) -> object: 
        """
        Return the result of the Future, blocking if the Future is not yet done.

        If the Future has a result, its value is returned.

        If the Future has an exception, raises the exception.
        """
    pass
class _PyObjectFutureIter():
    def __iter__(self) -> object: ...
    def __next__(self) -> None: ...
    pass
class _Read_Array_bool_1d(IPyFromExcel):
    def __init__(self, trim: bool = True) -> None: ...
    pass
class _Read_Array_bool_2d(IPyFromExcel):
    def __init__(self, trim: bool = True) -> None: ...
    pass
class _Read_Array_datetime_1d(IPyFromExcel):
    def __init__(self, trim: bool = True) -> None: ...
    pass
class _Read_Array_datetime_2d(IPyFromExcel):
    def __init__(self, trim: bool = True) -> None: ...
    pass
class _Read_Array_datetime_1d():
    pass
class _Read_Array_datetime_2d():
    pass
class _Read_Array_float_1d(IPyFromExcel):
    def __init__(self, trim: bool = True) -> None: ...
    pass
class _Read_Array_float_2d(IPyFromExcel):
    def __init__(self, trim: bool = True) -> None: ...
    pass
class _Read_Array_int_1d(IPyFromExcel):
    def __init__(self, trim: bool = True) -> None: ...
    pass
class _Read_Array_int_2d(IPyFromExcel):
    def __init__(self, trim: bool = True) -> None: ...
    pass
class _Read_Array_object_1d(IPyFromExcel):
    def __init__(self, trim: bool = True) -> None: ...
    pass
class _Read_Array_object_2d(IPyFromExcel):
    def __init__(self, trim: bool = True) -> None: ...
    pass
class _Read_Array_str_1d(IPyFromExcel):
    def __init__(self, trim: bool = True) -> None: ...
    pass
class _Read_Array_str_2d(IPyFromExcel):
    def __init__(self, trim: bool = True) -> None: ...
    pass
class _Read_Cache(IPyFromExcel):
    def __init__(self) -> None: ...
    pass
class _Read_Range(IPyFromExcel):
    def __init__(self) -> None: ...
    pass
class _Read__Uncached_bool(IPyFromExcel):
    def __init__(self) -> None: ...
    pass
class _Read__Uncached_float(IPyFromExcel):
    def __init__(self) -> None: ...
    pass
class _Read__Uncached_int(IPyFromExcel):
    def __init__(self) -> None: ...
    pass
class _Read__Uncached_object(IPyFromExcel):
    def __init__(self) -> None: ...
    pass
class _Read__Uncached_str(IPyFromExcel):
    def __init__(self) -> None: ...
    pass
class _Read_bool(IPyFromExcel):
    def __init__(self) -> None: ...
    pass
class _Read_date(IPyFromExcel):
    def __init__(self) -> None: ...
    pass
class _Read_datetime(IPyFromExcel):
    def __init__(self) -> None: ...
    pass
class _Read_dict(IPyFromExcel):
    def __init__(self) -> None: ...
    pass
class _Read_float(IPyFromExcel):
    def __init__(self) -> None: ...
    pass
class _Read_int(IPyFromExcel):
    def __init__(self) -> None: ...
    pass
class _Read_object(IPyFromExcel):
    def __init__(self) -> None: ...
    pass
class _Read_str(IPyFromExcel):
    def __init__(self) -> None: ...
    pass
class _Read_tuple_from_Excel(IPyFromExcel):
    def __init__(self) -> None: ...
    pass
class _Return_Array_bool_1d(IPyToExcel):
    def __init__(self, cache: bool = False) -> None: ...
    pass
class _Return_Array_bool_2d(IPyToExcel):
    def __init__(self, cache: bool = False) -> None: ...
    pass
class _Return_Array_datetime_1d(IPyToExcel):
    def __init__(self, cache: bool = False) -> None: ...
    pass
class _Return_Array_datetime_2d(IPyToExcel):
    def __init__(self, cache: bool = False) -> None: ...
    pass
class _Return_Array_datetime_1d():
    pass
class _Return_Array_datetime_2d():
    pass
class _Return_Array_float_1d(IPyToExcel):
    def __init__(self, cache: bool = False) -> None: ...
    pass
class _Return_Array_float_2d(IPyToExcel):
    def __init__(self, cache: bool = False) -> None: ...
    pass
class _Return_Array_int_1d(IPyToExcel):
    def __init__(self, cache: bool = False) -> None: ...
    pass
class _Return_Array_int_2d(IPyToExcel):
    def __init__(self, cache: bool = False) -> None: ...
    pass
class _Return_Array_object_1d(IPyToExcel):
    def __init__(self, cache: bool = False) -> None: ...
    pass
class _Return_Array_object_2d(IPyToExcel):
    def __init__(self, cache: bool = False) -> None: ...
    pass
class _Return_Array_str_1d(IPyToExcel):
    def __init__(self, cache: bool = False) -> None: ...
    pass
class _Return_Array_str_2d(IPyToExcel):
    def __init__(self, cache: bool = False) -> None: ...
    pass
class _Return_Cache(IPyToExcel):
    def __init__(self) -> None: ...
    pass
class _Return_SingleValue(IPyToExcel):
    def __init__(self) -> None: ...
    pass
class _Return_bool(IPyToExcel):
    def __init__(self) -> None: ...
    pass
class _Return_date(IPyToExcel):
    def __init__(self) -> None: ...
    pass
class _Return_datetime(IPyToExcel):
    def __init__(self) -> None: ...
    pass
class _Return_dict(IPyToExcel):
    def __init__(self) -> None: ...
    pass
class _Return_float(IPyToExcel):
    def __init__(self) -> None: ...
    pass
class _Return_int(IPyToExcel):
    def __init__(self) -> None: ...
    pass
class _Return_str(IPyToExcel):
    def __init__(self) -> None: ...
    pass
class _Return_tuple_to_Excel(IPyToExcel):
    def __init__(self) -> None: ...
    pass
def _get_event_loop(arg0: str) -> None:
    pass
def active_workbook() -> Workbook:
    """
    Returns the currently active workbook. Will raise an exception if xlOil
    has not been loaded as an addin.
    """
def active_worksheet() -> Worksheet:
    """
    Returns the currently active worksheet. Will raise an exception if xlOil
    has not been loaded as an addin.
    """
def app() -> Application:
    """
    Returns the parent Excel Application object when xlOil is loaded as an
    addin. Will throw if xlOil has been imported to run automation.
    """
def call(func: object, *args) -> object:
    """
    Calls a built-in worksheet function or command or a user-defined function with the 
    given name. The name is case-insensitive; built-in functions take priority in a name
    clash.

    The type and order of arguments expected depends on the function being called.  

    `func` can be built-in function number (as an int) which slightly reduces the lookup overhead

    This function must be called from a *non-local worksheet function on the main thread*.

    `call` can also invoke old-style `macro sheet commands <https://docs.excel-dna.net/assets/excel-c-api-excel-4-macro-reference.pdf>`_
    """
def call_async(func: object, *args) -> _ExcelObjFuture:
    """
    Calls a built-in worksheet function or command or a user-defined function with the 
    given name.  See ``xloil.call``.

    Calls to the Excel API must be done on Excel's main thread: this async function
    can be called from any thread but will require the main thread to be available
    to return a result.

    Returns an **awaitable**, i.e. a future which holds the result.
    """
def create_gui(ribbon: object = None, func_names: object = None, name: object = None) -> _AddinFuture:
    """
    Returns an **awaitable** to a ExcelGUI object which passes the specified ribbon
    customisation XML to Excel.  When the returned object is deleted, it 
    unloads the Ribbon customisation and the associated COM add-in.  If ribbon
    XML is specfied the ExcelGUI object will be connected, otherwise the 
    user must call the `connect()` method to active the object.

    Parameters
    ----------

    ribbon: str
        A Ribbon XML string, most easily created with a specialised editor.
        The XML format is documented on Microsoft's website

    func_names: Func[str -> callable] or Dict[str, callabe]
        The ``func_names`` mapper links callbacks named in the Ribbon XML to
        python functions. It can be either a dictionary containing named 
        functions or any callable which returns a function given a string.
        Each return handler should take a single ``RibbonControl``
        argument which describes the control which raised the callback.

        Callbacks declared async will be executed in the addin's event loop. 
        Other callbacks are executed in Excel's main thread. Async callbacks 
        cannot return values.

    name: str
        The addin name which will appear in Excel's COM addin list.
        If None, uses the filename at the call site as the addin name.
    """
def deregister_functions(arg0: object, arg1: object) -> None:
    """
    Deregisters worksheet functions linked to specified module. Generally, there
    is no need to call this directly.
    """
def excel_callback(func: object, wait: int = 0, retry: int = 500, api: str = '') -> _PyObjectFuture:
    """
    Schedules a callback to be run in the main thread. Much of the COM API in unavailable
    during the calc cycle, in particular anything which involves writing to the sheet.
    Returns a future which can be awaited.

    Parameters
    ----------

    func: callable
    A callable which takes no arguments and returns nothing

    retry : int
    Millisecond delay between retries if Excel's COM API is busy, e.g. a dialog box
    is open or it is running a calc cycle.If zero, does no retry

    wait : int
    Number of milliseconds to wait before first attempting to run this function

    api : str
    Specify 'xll' or 'com' or both to indicate which APIs the call requires.
    The default is 'com': 'xll' would only be required in rare cases.
    """
def excel_state() -> ExcelState:
    """
    Gives information about the Excel application, in particular the handles required
    to interact with Excel via the Win32 API. Only available when xlOil is loaded as 
    an addin.
    """
def from_excel_date(arg0: object) -> object:
    """
    Tries to the convert a given number to a `dt.date` or `dt.datetime` assuming it is an 
    Excel date serial number.  Strings are parsed using the current date conversion 
    settings. If `dt.datetime` is provided, it is simply returned as is.  Raises `ValueError`
    if conversion is not possible.
    """
def get_async_loop() -> object:
    """
    Returns the asyncio event loop associated with the async background
    worker thread.  All async / RTD worksheet functions are executed 
    on this event loop.
    """
def in_wizard() -> bool:
    """
    Returns true if the function is being invoked from the function wizard : costly functions should"
    exit in this case to maintain UI responsiveness.Checking for the wizard is itself not cheap, so"
    use this sparingly.
    """
def insert_cell_image(writer: object, size: object = None, pos: object = None, origin: object = None, compress: bool = True) -> str:
    """
    Inserts an image associated with the calling cell. A second call to this function
    removes any image previously inserted from the same calling cell.

    Parameters
    ----------

    writer: 
        a one-arg function which writes the image to a provided filename. The file
        format must be one that Excel can open.
    size:  
        * A tuple (width, height) in points. 
        * "cell" to fit to the caller size
        * "img" or None to keep the original image size
    pos:
        A tuple (X, Y) in points. The origin is determined by the `origin` argument
    origin:
        * "top" or None: the top left of the calling range
        * "sheet": the top left of the sheet
        * "bottom": the bottom right of the calling range
    compress:
        if True, compresses the resulting image before storing in the sheet
    """
def register_functions(funcs: typing.List[_FuncSpec], module: object = None, addin: object = None, append: bool = False) -> None:
    pass
def run(func: object, *args) -> object:
    """
    Calls VBA's `Application.Run` taking the function name and up to 30 arguments.
    This can call any user-defined function or macro but not built-in functions.

    The type and order of arguments expected depends on the function being called.

    Must be called on Excel's main thread, for example in worksheet function or 
    command.
    """
def run_async(func: object, *args) -> _ExcelObjFuture:
    """
    Calls VBA's `Application.Run` taking the function name and up to 30 arguments.
    This can call any user-defined function or macro but not built-in functions.

    Calls to the Excel API must be done on Excel's main thread: this async function
    can be called from any thread but will require the main thread to be available
    to return a result.

    Returns an **awaitable**, i.e. a future which holds the result.
    """
_return_converter_hook: xloil_core._CustomReturnConverter = None
cache: xloil_core.ObjectCache = None

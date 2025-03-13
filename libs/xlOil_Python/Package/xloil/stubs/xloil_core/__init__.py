"""
        The Python plugin for xlOil primarily allows creation of Excel functions and macros 
        backed by Python code. In addition it offers full control over GUI objects and an 
        interface for Excel automation: driving the application in code.

        See the documentation at https://xloil.readthedocs.io
      """
from __future__ import annotations
import typing

__all__ = [
    "Addin",
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
    "active_cell",
    "active_workbook",
    "active_worksheet",
    "all_workbooks",
    "app",
    "cache",
    "call",
    "call_async",
    "core_addin",
    "date_formats",
    "deregister_functions",
    "event",
    "excel_callback",
    "excel_state",
    "from_excel_date",
    "get_async_loop",
    "in_wizard",
    "insert_cell_image",
    "run",
    "run_async",
    "selection",
    "to_datetime",
    "xloil_addins"
]


class Addin():
    def __repr__(self) -> str: ...
    def __str__(self) -> str: ...
    def functions(self) -> typing.List[_FuncSpec]: 
        """
        Returns a list of all functions declared by this addin.
        """
    def source_files(self) -> typing.List[str]: ...
    @property
    def async_slice(self) -> int:
        """
                      Sets/gets the time slice in milliseconds for which the asyncio event loop is allowed 
                      to run before being interrupted. The event loop holds the GIL while it is running, so
                      making this interval too long could impact the performance of other python functions.
                    

        :type: int
        """
    @async_slice.setter
    def async_slice(self, arg1: int) -> None:
        """
        Sets/gets the time slice in milliseconds for which the asyncio event loop is allowed 
        to run before being interrupted. The event loop holds the GIL while it is running, so
        making this interval too long could impact the performance of other python functions.
        """
    @property
    def async_throttle(self) -> int:
        """
                      Sets/gets the interval in milliseconds between switches to the asyncio event loop
                      embedded in this addin. The event loop holds the GIL while it is running, so making
                      this interval too short could impact the performance of other python functions.
                    

        :type: int
        """
    @async_throttle.setter
    def async_throttle(self, arg1: int) -> None:
        """
        Sets/gets the interval in milliseconds between switches to the asyncio event loop
        embedded in this addin. The event loop holds the GIL while it is running, so making
        this interval too short could impact the performance of other python functions.
        """
    @property
    def event_loop(self) -> object:
        """
                      The asyncio event loop used for background tasks by this addin
                    

        :type: object
        """
    @property
    def pathname(self) -> str:
        """
        :type: str
        """
    @property
    def settings(self) -> object:
        """
                      Gives access to the settings in the addin's ini file as nested dictionaries.
                      These are the settings on load and do not allow for modifications made in the 
                      ribbon toolbar.
                    

        :type: object
        """
    @property
    def settings_file(self) -> str:
        """
                      The full pathname of the settings ini file used by this addin
                    

        :type: str
        """
    pass
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

    ::

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
    def __setattr__(self, arg0: object, arg1: object) -> None: ...
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
    def open(self, filepath: str, update_links: bool = True, read_only: bool = False, delimiter: object = None) -> Workbook: 
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
    def run(self, func: str, *args) -> object: 
        """
        Calls VBA's `Application.Run` taking the function name and up to 30 arguments.
        This can call any user-defined function or macro but not built-in functions.

        The type and order of arguments expected depends on the function being called.
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
    def active_cell(self) -> object:
        """
                      Returns the currently active cell as a Range or None.
                  

        :type: object
        """
    @property
    def active_workbook(self) -> object:
        """
                      Returns the currently active workbook or None.
                  

        :type: object
        """
    @property
    def active_worksheet(self) -> object:
        """
                      Returns the currently active worksheet or None.
                  

        :type: object
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
    def has_dynamic_arrays(self) -> bool:
        """
        :type: bool
        """
    @property
    def selection(self) -> object:
        """
                      Returns the currently active cell as a Range or None.
                  

        :type: object
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
    def workbook_paths(self) -> None:
        """
        A set of the full path names of all workbooks open in this Application. Does not use COM interface.

        :type: None
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
        Range object corresponding to caller address.  Will raise an exception if caller is not a range

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
                  registered function argument does not explicitly specify a type 
                  (e.g. int or str via an annotation), it may be passed a *CellError*, 
                  which it can handle based on the error type.

                  The integer value of a *CellError* corresponds to it's VBA/COM error
                  number, so for example we can write 
                  `if cell.Value2 == CellError.NA.value: ...`
                

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
    DIV: xloil_core.CellError=None # value = <CellError.DIV: -2146826281>
    GETTING_DATA: xloil_core.CellError=None # value = <CellError.GETTING_DATA: -2146826245>
    NA: xloil_core.CellError=None # value = <CellError.NA: -2146826246>
    NAME: xloil_core.CellError=None # value = <CellError.NAME: -2146826259>
    NULL: xloil_core.CellError=None # value = <CellError.NULL: -2146826288>
    NUM: xloil_core.CellError=None # value = <CellError.NUM: -2146826252>
    REF: xloil_core.CellError=None # value = <CellError.REF: -2146826265>
    VALUE: xloil_core.CellError=None # value = <CellError.VALUE: -2146826273>
    __members__: dict=None # value = {'NULL': <CellError.NULL: -2146826288>, 'DIV': <CellError.DIV: -2146826281>, 'VALUE': <CellError.VALUE: -2146826273>, 'REF': <CellError.REF: -2146826265>, 'NAME': <CellError.NAME: -2146826259>, 'NUM': <CellError.NUM: -2146826252>, 'NA': <CellError.NA: -2146826246>, 'GETTING_DATA': <CellError.GETTING_DATA: -2146826245>}
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
    An `ExcelGUI` wraps a COM addin which allows Ribbon customisation and creation
    of custom task panes. 

    The methods of this object are safe to call from any thread; however, since COM calls 
    must be made on Excel's main thread, the methods schedule  those calls and return an 
    *awaitable* future to the result. This could lead to deadlocks if, for example, the 
    future's result is requested synchronously and one of Excel's event handlers is 
    triggered.  For a safer non-blocking approach, use `excel_callback` to invoke code 
    which manipulates the *ExcelGUI* object.

    The object's properties do not return futures and are thread-safe.
    """
    def __init__(self, name: object = None, ribbon: object = None, funcmap: object = None, connect: bool = True) -> None: 
        """
        Creates an `ExcelGUI` using the specified ribbon customisation XML
        and optionally connects it to Excel, ready for use.

        When the *ExcelGUI* object is deleted, it unloads the associated COM 
        add-in and so all Ribbon customisation and attached task panes.

        Parameters
        ----------

        ribbon: str
            A Ribbon XML string, most easily created with a specialised editor.
            The XML format is documented on Microsoft's website

        funcmap: Func[str -> callable] or Dict[str, callabe]
            The ``funcmap`` mapper links callbacks named in the Ribbon XML to
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

        connect: bool
            If True (the default) initiates a connection request for the addin, 
            which causes Excel to load it and any associated ribbons. The connection
            is *scheduled* in Excel's main thread when the COM interface is available.
            Note that COM is not available when xlOil is starting up.

            If you want to call other methods on *ExcelGUI* then to ensure it is 
            connected use `ExcelGUI.connect().result()` to block on the future or
            use `excel_callback`.
        """
    def _create_task_pane_frame(self, name: str, progid: object = None, window: object = None) -> _CTPFuture: 
        """
        Used internally to create a custom task pane window which can be populated
        with a python GUI.  Most users should use `attach_pane(...)` instead.

        A COM `progid` can be specified, but this will prevent displaying a python GUI
        in the task pane using the xlOil methods. This is a specialised use case.
        """
    def activate(self, id: str) -> _Future: 
        """
        Activatives the ribbon tab with the specified id.  Returns False if
        there is no Ribbon or the Ribbon is collapsed.
        """
    def attach_pane(self, arg0: object, arg1: object, arg2: object, arg3: object, arg4: object) -> object: 
        """
        Given task pane contents (which can be specified in several forms) this function
        creates a new task pane displaying those contents.

        Returns the instance of `CustomTaskPane`.  If one was passed as the 'pane' argument, 
        that is returned, if a *QWidget* was passed, a `QtThreadTaskPane` is created.

        Parameters
        ----------

        pane: CustomTaskPane (or QWidget type)
            Can be an instance of `CustomTaskPane`, a type deriving from `QWidget` or
            an instance of a `QWidget`. If a QWidget instance is passed, it must have 
            been created on the Qt thread.

        name: 
            The task pane name. Will be displayed above the task pane. If not provided,
            the 'name' attribute of the task pane is used.

        window: 
            A window title or `xloil.ExcelWindow` object to which the task pane should be
            attached.  If None, the active window is used.

        size:
            If provided, a tuple (width, height) used to set the initial pane size

        visible:
            Determines the initial pane visibility. Defaults to True.
        """
    def attach_pane_async(self, pane: object, name: object = None, window: object = None, size: object = None, visible: object = True) -> object: 
        """
        Behaves as per `attach_pane`, but returns an *asyncio* coroutine. The
        `pane` argument may be an awaitable to a `CustomTaskPane`.
        """
    def connect(self) -> _Future: 
        """
        Connects the underlying COM addin to Excel, No other methods may be called 
        on a `ExcelGUI` object until it has been connected.

        This method is safe to call on an already-connected addin.
        """
    def create_task_pane(self, name: object, creator: object, window: object = None, size: object = None, visible: object = True) -> object: 
        """
        Deprecated: use `attach_pane`. Note that `create_task_pane` tries to `find_task_pane`
        before creation whereas `attach_pane` does not.
        """
    def disconnect(self) -> _Future: 
        """
        Unloads the underlying COM add-in and any ribbon customisation.  Avoid using
        connect/disconnect to modify the Ribbon as it is not perfomant. Rather hide/show
        controls with `invalidate` and the vibility callback.
        """
    def invalidate(self, id: str = '') -> _Future: 
        """
        Invalidates the specified control: this clears the cache of responses
        to callbacks associated with the control. For example, this can be
        used to hide a control by forcing its getVisible callback to be invoked,
        rather than using the cached value.

        If no control ID is specified, all controls are invalidated.
        """
    @property
    def connected(self) -> bool:
        """
        True if the a connection to Excel has been made

        :type: bool
        """
    @property
    def name(self) -> str:
        """
        The name displayed in Excel's COM Addins window

        :type: str
        """
    pass
class ExcelState():
    """
    Gives information about the Excel application. Cannot be constructed: call
    ``xloil.excel_state`` to get an instance.
    """
    @property
    def hinstance(self) -> int:
        """
        Excel Win32 HINSTANCE pointer as an int

        :type: int
        """
    @property
    def hwnd(self) -> int:
        """
        Excel Win32 main window handle as an int

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
    A document window which displays a view of a workbook.
    See `Excel.Window <https://docs.microsoft.com/en-us/office/vba/api/excel.WindowWindow>`_ 
    """
    def __getattr__(self, arg0: str) -> object: ...
    def __setattr__(self, arg0: object, arg1: object) -> None: ...
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
    A collection of all the document window objects in Excel. A document window 
    shows a view of a Workbook.

    See `Excel.Windows <https://docs.microsoft.com/en-us/office/vba/api/excel.WindowsWindows>`_ 
    """
    def __contains__(self, arg0: str) -> bool: ...
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
    def __str__(self) -> str: ...
    pass
class IPyToExcel():
    def __call__(self, arg0: _object) -> _RawExcelValue: ...
    def __str__(self) -> str: ...
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
        Adds an object to the cache and returns a reference string.

        xlOil automatically adds objects returned from worksheet 
        functions to the cache if they cannot be converted by any 
        registered converter.  So this function is useful to:

           1) force a convertible object, such as an iterable, into the
              cache
           2) return a list of cached objects
           3) create cached objects from outside of worksheet fnctions
              e.g. in commands / subroutines

        xlOil uses the caller infomation provided by Excel to construct
        the cache string and manage the cache object lifecycle. When
        invoked from a worksheet function, this caller info contains 
        the cell reference. xlOil deletes cache objects linked to the 
        cell reference from previous calculation cycles.

        When invoked from a source other than a worksheet function (there
        are several possibilies, see the help for `xlfCaller`), xlOil
        again generates a reference string based on the caller info. 
        However, this may not be unique.  In addition, objects with the 
        same caller string will replace those created during a previous 
        calculation cycle. For example, creating cache objects from a button
        clicked repeatedly will behave differently if Excel recalculates 
        in between the clicks. To override this behaviour, the exact cache
        `key` can be specified.  For example, use Python's `id` function or
        the cell address being written to if a command is writing a cache
        string to the sheet.  When `key` is specified the user is responsible
        for managing the lifecycle of their cache objects.


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
    def remove(self, ref: str) -> bool: 
        """
        xlOil manages the lifecycle for most cache objects, so this  
        function should only be called when `add` was invoked with a
        specified key - in this case the user owns the lifecycle 
        management. 
        """
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

        x[1, 1] # The *value* at (1, 1) as a python type: int, str, float, etc.

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
    def __iadd__(self, arg0: object) -> object: ...
    def __imul__(self, arg0: object) -> object: ...
    def __init__(self, address: str) -> None: ...
    def __isub__(self, arg0: object) -> object: ...
    def __iter__(self) -> RangeIter: ...
    def __itruediv__(self, arg0: object) -> object: ...
    def __len__(self) -> int: ...
    def __setattr__(self, arg0: object, arg1: object) -> None: ...
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
    def offset(self, from_row: int, from_col: int, num_rows: object = None, num_cols: object = None) -> Range: 
        """
        Similar to the *range* function, but with different defaults  

        Parameters
        ----------

        from_row: int
            Starting row offset from the top left of the parent range. Zero-based, can be negative

        from_col: int
            Starting row offset from the top left of the parent range. Zero-based, can be negative

        num_rows: int
            Number of rows in output range. Defaults to 1

        num_cols: int
            Number of columns in output range. Defaults to 1.
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
    def set_formula(self, formula: object, how: str = '') -> None: 
        """
        The `how` parameter determines how this function differs from setting the `formula` 
        property:

          * *dynamic* (or omitted): identical to setting the `formula` property
          * *array*: if the target range is larger than one cell and a single string is passed,
            set this as an array formula for the range
          * *implicit*: uses old-style implicit intersection - see "Formula vs Formula2" on MSDN
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
    def formula(self) -> object:
        """
                    Get / sets the formula for the range. If the cell contains a constant, this property returns 
                    the value. If the cell is empty, this property returns an empty string. If the cell contains
                    a formula, the property returns the formula that would be displayed in the formula bar as a
                    string.  If the range is larger than one cell, the property returns an array of the values  
                    which would be obtained calling `formula` on each cell.
                    
                    When setting, if the range is larger than one cell and a single value is passed that value
                    is filled into each cell. Alternatively, you can set the formula to an array of the same 
                    dimensions.
                  

        :type: object
        """
    @formula.setter
    def formula(self, arg1: object) -> None:
        """
        Get / sets the formula for the range. If the cell contains a constant, this property returns 
        the value. If the cell is empty, this property returns an empty string. If the cell contains
        a formula, the property returns the formula that would be displayed in the formula bar as a
        string.  If the range is larger than one cell, the property returns an array of the values  
        which would be obtained calling `formula` on each cell.

        When setting, if the range is larger than one cell and a single value is passed that value
        is filled into each cell. Alternatively, you can set the formula to an array of the same 
        dimensions.
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
    def set_done(self) -> None: 
        """
        Indicates that the task has completed and the RtdReturn can drop its reference
        to the task. Further calls to `set_result()` will be ignored.
        """
    def set_result(self, arg0: object) -> None: ...
    def set_task(self, task: object) -> None: 
        """
        Set the task object to keep it alive until the task indicates it is done. The
        task object should respond to the `cancel()` method.
        """
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
        Registers an RtdPublisher with this manager. The RtdPublisher receives
        notification when the number of subscribers changes
        """
    def start_task(self, topic: str, func: object, converter: IPyToExcel = None) -> None: 
        """
        Launch a publishing task for a `topic` given a func and a return converter.
        The function should take a single `xloil.RtdReturn` argument.
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
    @property
    def progid(self) -> str:
        """
        :type: str
        """
    pass
class StatusBar():
    """
     
    Displays status bar messages and clears the status bar (after an optional delay) 
    on context exit.

    Examples
    --------

    ::

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
    Manages Excel's underlying custom task pane object into which a python GUI can be
    drawn. It is unlikely that this object will need to be manipulated directly. Rather 
    use `xloil.gui.CustomTaskPane` which holds the python-side frame contents.

    The methods of this object are safe to call from any thread. COM must be used on 
    Excel's main thread, so the methods all wrap their calls to ensure to this happens.
    """
    def attach(self, handler: object, hwnd: int, as_parent: bool = True) -> _Future: 
        """
        Associates a `xloil.gui.CustomTaskPane` with this frame. Returns a future
        with no result.
        """
    def com_control(self, lib: str = '') -> object: 
        """
        Gets the base COM control of the task pane. The ``lib`` used to provide
        COM support can be 'comtypes' or 'win32com' (default is win32com). This 
        method is only useful if a custom `progid` was specified during the task
        pane creation.
        """
    @property
    def position(self) -> str:
        """
                      Gets/sets the dock position, one of: bottom, floating, left, right, top
                    

        :type: str
        """
    @position.setter
    def position(self, arg1: str) -> None:
        """
        Gets/sets the dock position, one of: bottom, floating, left, right, top
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
    A handle to an open Excel workbook.
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
    def __setattr__(self, arg0: object, arg1: object) -> None: ...
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
    def __contains__(self, arg0: str) -> bool: ...
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
        equivalent to ``at``.  

        If the argument is a 2-tuple, slices the sheet to return an xloil.Range.
        Uses normal python slicing conventions, i.e [left included, right excluded), negative
        numbers are offset from the end.
        """
    def __setattr__(self, arg0: object, arg1: object) -> None: ...
    def __setitem__(self, arg0: object, arg1: object) -> None: ...
    def __str__(self) -> str: ...
    def activate(self) -> None: 
        """
        Makes this worksheet the active sheet
        """
    def at(self, address: str) -> _ExcelRange: 
        """
        Returns the range specified by the local address, e.g. ``.at('B3:D6')``
        """
    def calculate(self) -> None: 
        """
        Calculates this worksheet
        """
    def cell(self, row: int, col: int) -> _ExcelRange: 
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
    @property
    def used_range(self) -> _ExcelRange:
        """
        Returns a Range object that represents the used range on the worksheet

        :type: _ExcelRange
        """
    pass
class Worksheets():
    """
    A collection of all the Worksheet objects in the specified or active workbook. 

    See `Excel.Worksheets <https://docs.microsoft.com/en-us/office/vba/api/excel.WorksheetsWorksheets>`_ 
    """
    def __contains__(self, arg0: str) -> bool: ...
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
class _AddinsDict():
    """
    A dictionary of all addins using the xlOil_Python plugin keyed
    by the addin pathname.
    """
    def __contains__(self, arg0: str) -> bool: ...
    def __getitem__(self, arg0: str) -> Addin: ...
    def __iter__(self) -> typing.Iterator: ...
    def __len__(self) -> int: ...
    def items(self) -> typing.Iterator: ...
    def keys(self) -> typing.Iterator: ...
    def values(self) -> typing.Iterator: ...
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
    def __init__(self, callable: object, check_cache: bool = True, name: str = 'custom') -> None: ...
    pass
class _CustomReturn(IPyToExcel):
    def __init__(self, callable: object, name: str = 'custom') -> None: ...
    def invoke(self, arg0: object) -> object: ...
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
class _DateFormatList():
    """
    Registers date time formats to try when parsing strings to dates.
    See `std::get_time` for format syntax.
    """
    def __bool__(self) -> bool: 
        """
        Check whether the list is nonempty
        """
    def __contains__(self, x: str) -> bool: 
        """
        Return true the container contains ``x``
        """
    @typing.overload
    def __delitem__(self, arg0: int) -> None: 
        """
        Delete the list elements at index ``i``

        Delete list elements using a slice object
        """
    @typing.overload
    def __delitem__(self, arg0: slice) -> None: ...
    def __eq__(self, arg0: _DateFormatList) -> bool: ...
    @typing.overload
    def __getitem__(self, arg0: int) -> str: 
        """
        Retrieve list elements using a slice object
        """
    @typing.overload
    def __getitem__(self, s: slice) -> _DateFormatList: ...
    @typing.overload
    def __init__(self) -> None: 
        """
        Copy constructor
        """
    @typing.overload
    def __init__(self, arg0: _DateFormatList) -> None: ...
    @typing.overload
    def __init__(self, arg0: typing.Iterable) -> None: ...
    def __iter__(self) -> typing.Iterator: ...
    def __len__(self) -> int: ...
    def __ne__(self, arg0: _DateFormatList) -> bool: ...
    @typing.overload
    def __setitem__(self, arg0: int, arg1: str) -> None: 
        """
        Assign list elements using a slice object
        """
    @typing.overload
    def __setitem__(self, arg0: slice, arg1: _DateFormatList) -> None: ...
    def append(self, x: str) -> None: 
        """
        Add an item to the end of the list
        """
    def clear(self) -> None: 
        """
        Clear the contents
        """
    def count(self, x: str) -> int: 
        """
        Return the number of times ``x`` appears in the list
        """
    @typing.overload
    def extend(self, L: _DateFormatList) -> None: 
        """
        Extend the list by appending all the items in the given list

        Extend the list by appending all the items in the given list
        """
    @typing.overload
    def extend(self, L: typing.Iterable) -> None: ...
    def insert(self, i: int, x: str) -> None: 
        """
        Insert an item at a given position.
        """
    @typing.overload
    def pop(self) -> str: 
        """
        Remove and return the last item

        Remove and return the item at index ``i``
        """
    @typing.overload
    def pop(self, i: int) -> str: ...
    def remove(self, x: str) -> None: 
        """
        Remove the first item from the list whose value is x. It is an error if there is no such item.
        """
    __hash__ = None
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
class _ExcelRange(Range):
    pass
class _FuncArg():
    def __init__(self, arg0: str, arg1: str, arg2: IPyFromExcel, arg3: str) -> None: ...
    def __str__(self) -> str: ...
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
    def flags(self) -> str:
        """
        :type: str
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
    pass
class _FuncSpec():
    def __init__(self, func: function, args: typing.List[_FuncArg], name: str = '', features: str = None, help: str = '', category: str = '', local: bool = True, volatile: bool = False, errors: int = 0) -> None: ...
    def __str__(self) -> str: ...
    @property
    def args(self) -> typing.List[_FuncArg]:
        """
        :type: typing.List[_FuncArg]
        """
    @property
    def error_propagation(self) -> bool:
        """
                      Used internally to control the error propagation setting
                    

        :type: bool
        """
    @error_propagation.setter
    def error_propagation(self, arg1: bool) -> None:
        """
        Used internally to control the error propagation setting
        """
    @property
    def func(self) -> function:
        """
                      Yes you can change the function which is called by Excel! Use
                      with caution.
                    

        :type: function
        """
    @func.setter
    def func(self, arg1: function) -> None:
        """
        Yes you can change the function which is called by Excel! Use
        with caution.
        """
    @property
    def help(self) -> str:
        """
        :type: str
        """
    @property
    def is_async(self) -> bool:
        """
                      True if the function used Excel's native async
                    

        :type: bool
        """
    @property
    def is_rtd(self) -> bool:
        """
                      True if the function uses RTD to provide async returns
                    

        :type: bool
        """
    @property
    def is_threaded(self) -> bool:
        """
                      True if the function can be multi-threaded during Excel calcs
                    

        :type: bool
        """
    @property
    def name(self) -> str:
        """
                      Writing to name property doesn't make sense when registered
                    

        :type: str
        """
    @name.setter
    def name(self, arg1: str) -> None:
        """
        Writing to name property doesn't make sense when registered
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
    def __call__(self, msg: object, *args, **kwargs) -> None: 
        """
        Writes a message to the log at the specifed keyword paramter `level`. The default 
        level is 'info'.  The message can contain format specifiers which are expanded
        using any additional positional arguments. This allows for lazy contruction of the 
        log string like python's own 'logging' module.
        """
    def __init__(self) -> None: 
        """
        Do not construct this class - a singleton instance is created by xlOil.
        """
    def debug(self, msg: object, *args) -> None: 
        """
        Writes a log message at the 'debug' level
        """
    def error(self, msg: object, *args) -> None: 
        """
        Writes a log message at the 'error' level
        """
    def flush(self) -> None: 
        """
        Forces a log file 'flush', i.e write pending log messages to the log file.
        For performance reasons the file is not by default flushed for every message.
        """
    def info(self, msg: object, *args) -> None: 
        """
        Writes a log message at the 'info' level
        """
    def trace(self, msg: object, *args) -> None: 
        """
        Writes a log message at the 'trace' level
        """
    def warn(self, msg: object, *args) -> None: 
        """
        Writes a log message at the 'warn' level
        """
    @property
    def flush_on(self) -> str:
        """
                      Returns or sets the log level which will trigger a 'flush', i.e a writing pending
                      log messages to the log file.
                    

        :type: str
        """
    @flush_on.setter
    def flush_on(self, arg1: object) -> None:
        """
        Returns or sets the log level which will trigger a 'flush', i.e a writing pending
        log messages to the log file.
        """
    @property
    def level(self) -> str:
        """
                      Returns or sets the current log level. The returned value will always be an 
                      integer corresponding to levels in the `logging` module.  The level can be
                      set to an integer or one of the strings *error*, *warn*, *info*, *debug* or *trace*.
                    

        :type: str
        """
    @level.setter
    def level(self, arg1: object) -> None:
        """
        Returns or sets the current log level. The returned value will always be an 
        integer corresponding to levels in the `logging` module.  The level can be
        set to an integer or one of the strings *error*, *warn*, *info*, *debug* or *trace*.
        """
    @property
    def level_int(self) -> int:
        """
                      Returns the log level as an integer corresponding to levels in the `logging` module.
                      Useful if you want to condition some output based on being above a certain log
                      level.
                    

        :type: int
        """
    @property
    def levels(self) -> typing.List[str]:
        """
        A list of the available log levels

        :type: typing.List[str]
        """
    @property
    def path(self) -> str:
        """
        The full pathname of the log file

        :type: str
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
class _Read_list(IPyFromExcel):
    def __init__(self) -> None: ...
    pass
class _Read_object(IPyFromExcel):
    def __init__(self) -> None: ...
    pass
class _Read_str(IPyFromExcel):
    def __init__(self) -> None: ...
    pass
class _Read_tuple(IPyFromExcel):
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
class _Return_Single(IPyToExcel):
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
class _Return_tuple(IPyToExcel):
    def __init__(self) -> None: ...
    pass
class _Return_tuple():
    pass
class _Return_str(IPyToExcel):
    def __init__(self) -> None: ...
    pass
class _Return_tuple():
    pass
class _TomlTable():
    def __getitem__(self, arg0: str) -> object: ...
    pass
class _XllRange(Range):
    pass
def _get_onedrive_source(arg0: str) -> str:
    pass
def _register_functions(funcs: typing.List[_FuncSpec], module: object = None, addin: object = None, append: bool = False) -> None:
    pass
def _table_converter(n: int, m: int, columns: object = None, rows: object = None, headings: object = None, index: object = None, index_name: object = None, cache_objects: bool = False) -> _RawExcelValue:
    """
    For internal use. Converts a table like object (such as a pandas DataFrame) to 
    RawExcelValue suitable for returning to xlOil.
      
    n, m:
      the number of data fields and the length of the fields
    columns / rows: 
      a iterable of numpy array containing data, specified as columns 
      or rows (not both)
    headings:
      optional array of data field headings
    index:
      optional data field labels - one per data point
    index_name:
      optional headings for the index, should be a 1 dim iteratable of size
      num_index_levels * num_column_levels
    cache_objects:
      if True, place unconvertible objects in the cache and return a ref string
      if False, call str(x) on unconvertible objects
    """
def active_cell() -> object:
    """
    Returns the currently active cell as a Range or None. Will raise an exception if xlOil
    has not been loaded as an addin.
    """
def active_workbook() -> object:
    """
    Returns the currently active workbook or None. Will raise an exception if xlOil
    has not been loaded as an addin.
    """
def active_worksheet() -> object:
    """
    Returns the currently active worksheet or None. Will raise an exception if xlOil
    has not been loaded as an addin.
    """
def all_workbooks() -> Workbooks:
    """
    Collection of workbooks for the current application. Equivalent to 
    `xloil.app().workbooks`.
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
def core_addin() -> Addin:
    pass
def deregister_functions(funcs: object, module: object = None) -> None:
    """
    Deregisters worksheet functions linked to specified module. Generally, there
    is no need to call this directly.
    """
def excel_callback(func: object, wait: int = 0, retry: int = 500, api: str = '') -> _PyObjectFuture:
    """
    Schedules a callback to be run in the main thread. Much of the COM API in unavailable
    during the calc cycle, in particular anything which involves writing to the sheet.
    COM is also unavailable whilst xlOil is loading.

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
    Identical to `xloil.to_datetime`.
    """
def get_async_loop() -> object:
    """
    Returns the asyncio event loop associated with the async background
    worker thread.  All async / RTD worksheet functions are executed 
    on this event loop.
    """
def in_wizard() -> bool:
    """
    Returns true if the function is being invoked from the function wizard : costly functions 
    should exit in this case to maintain UI responsiveness.  Checking for the wizard is itself 
    not cheap, so use this sparingly.
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
def run(func: str, *args) -> object:
    """
    Calls VBA's `Application.Run` taking the function name and up to 30 arguments.
    This can call any user-defined function or macro but not built-in functions.

    The type and order of arguments expected depends on the function being called.

    Must be called on Excel's main thread, for example in worksheet function or 
    command.
    """
def run_async(func: str, *args) -> _ExcelObjFuture:
    """
    Calls VBA's `Application.Run` taking the function name and up to 30 arguments.
    This can call any user-defined function or macro but not built-in functions.

    Calls to the Excel API must be done on Excel's main thread: this async function
    can be called from any thread but will require the main thread to be available
    to return a result.

    Returns an **awaitable**, i.e. a future which holds the result.
    """
def selection() -> object:
    """
    Returns the currently selected cells as a Range or None. Will raise an exception if xlOil
    has not been loaded as an addin.
    """
def to_datetime(arg0: object) -> object:
    """
    Tries to the convert the given object to a `dt.date` or `dt.datetime`:

      * Numbers are assumed to be Excel date serial numbers. 
      * Strings are parsed using the current date conversion settings.
      * A numpy array of floats is treated as Excel date serial numbers and converted
        to n array of datetime64[ns].
      * `dt.datetime` is provided is simply returned.

    Raises `ValueError` if conversion is not possible.
    """
_return_converter_hook: _CustomReturnConverter=None # value = <xloil_core._CustomReturnConverter object>
cache: ObjectCache=None # value = <xloil_core.ObjectCache object>
date_formats: _DateFormatList=None # value = <xloil_core._DateFormatList object>
xloil_addins: _AddinsDict=None # value = <xloil_core._AddinsDict object>

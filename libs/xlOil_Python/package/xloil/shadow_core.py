
# TODO: how can we synchronise the help here with what you see when you actually import xloil_core

def in_wizard():
    """ 
    Returns true if the function is being invoked from the function wizard: costly functions should 
    exit in this case to maintain UI responsiveness. Checking for the wizard is itself not cheap, so 
    use this sparingly.
    """
    pass

def log(msg, level="info"):
    """
    Writes a log message at a level of *error*, *warn*, *info*, *debug* or *trace*.
    Only messages with a level higher than the log level defined in the xlOil
    settings will be output to the log file. Trace output can only be seen with 
    a debug build of xlOil.
    """
    pass

def run_later(func,
        num_retries = 10,
        retry_delay = 500,
        wait_time = 0):
    """
    Schedules a callback to be run in the COM API context. Much of the COM API in unavailable
    during the calc cycle, in particular anything which involves writing to the sheet.

    Parameters
    ----------

    func: callable
        A callable which takes no arguments and returns nothing

    num_retries: int
        Number of times to retry the call if Excel's COM API is busy, e.g. a dialog box
        is open or it is running a calc cycle

    retry_delay: int
        Millisecond delay between retries

    wait_time: int
        Number of milliseconds to wait before first attempting to run this function

    """
    pass

class _ExcelState:
    version = int()
    hinstance = int()
    hwnd = int()
    main_thread_id = int()

def get_excel_state() -> _ExcelState:
    """
    Gives information about the Excel application, in particular the handles required 
    to interact with Excel via the Win32 API. The function returns a class with the 
    following members:
        *version*:   Excel major version
        *hinstance*: Excel HINSTANCE
        *hwnd*:      Excel main window handle (as an int)
        *main_thread_id*: Excel's main thread ID
    """
    pass

class Range:
    """
    Similar to an Excel Range object, this class allows access to an area on a 
    worksheet. It uses similar syntax to Excel's object, supporting the ``cell``
    and ``range`` functions, however they are zero-based like python.

    A Range can be accessed and sliced using the usual syntax (the slice step must be 1):

    ::

        x[1, 1] # The value at (1, 1) as a python type: int, str, float, etc.

        x[1, :] # The second row as another Range object

        x[:-1, :-1] # A sub-range omitting the last row and column

    """
    def range(self, from_row, from_col, num_rows=None, num_cols=None, to_row=None, to_col=None):
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
            End row offset from the top left of the parent range. This row will be included in 
            the range. The offset is zero-based and can be negative to select ranges outside the
            parent range. Do not specify both `to_row` and `num_rows`.

        to_col: int
            End column offset from the top left of the parent range. This column will be included in 
            the range. The offset is zero-based and can be negative to select ranges outside the
            parent range. Do not specify both `to_col` and `num_cols`.

        num_rows: int
            Number of rows in output range. Must be positive. If neither `num_rows` or `to_rows` 
            are specified, the range ends at the last row of the parent range.

        num_cols: int
            Number of columns in output range. Must be positive. If neither `num_cols` or `to_cols` 
            are specified, the range ends at the last column of the parent range.
        """
        pass
    def cell(self, row, col):
        """ 
        Returns a Range object which consists of the single cell specified. Note the indices
        are zero-based from the top left of the parent range.
        """
        pass
    @property
    def value(self):
        """ 
        Property which gets or sets the value for a range. A fetched value is converted
        to the most appropriate Python type using the normal generic converter.

        If you use a horizontal array for the assignemnt, it is duplicated down to fill 
        the entire rectangle. If you use a vertical array, it is duplicated right to fill 
        the entire rectangle. If you use a rectangular array, and it is too small for the 
        rectangular range you want to put it in, that range is padded with #N/As.
        """
        pass
    def set(self, val):
        """
        Sets the data in the range to the provided value. If a single value is passed
        all cells will be set to the value. If a 2d-array is provided, the array will be
        pasted at the top-left of the range with the remainging cells being set to #N/A.
        If a 1d array is provided it will be pasted at the top left and repeated down or
        right depending on orientation.
        """
        pass
    def clear(self):
        """
        Sets all values in the range to the Nil/Empty type
        """
        pass
    def address(self,local=False):
        """
        Gets the range address in A1 format. The `local` parameter specifies whether
        the workbook and sheet name should be included. For example `local=True` gives
        "[Book1]Sheet1!F37" and `local=False` returns "F37".
        """
        pass
    @property
    def nrows(self):
        """ Returns the number of rows in the range """
        pass
    @property
    def ncols(self):
        """ Returns the number of columns in the range """
        pass
    def __getitem__(self, tuple):
        """ 
        Given a 2-tuple, slices the range to return a sub Range or a single element. Uses
        normal python slicing conventions.
        """
        pass

class ExcelArray:
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
    def __getitem__(self, tuple):
        """ 
        Given a 2-tuple, slices the array to return a sub ExcelArray or a 
        single element.
        """
        pass
    def to_numpy(self, dtype=None, dims=2):
        """
        Converts the array to a numpy array. If dtype is None, attempts to 
        discover one, otherwise raises an exception if values cannot be 
        converted to the specified dtype. dims can be 1 or 2
        """
        pass
    @property
    def dims(self):
        """ 
        Property which gives the dimension of the array: 1 or 2
        """
        pass
    @property
    def nrows(self):
        """ Returns the number of rows in the array """
        pass
    @property
    def ncols(self):
        """ Returns the number of columns in the array """
        pass

class CellError:
    """
    Enum-type class which represents an Excel error condition of the 
    form `#N/A!`, `#NAME!`, etc passed as a function argument. If a 
    function argument does not specify a type (e.g. int, str) it may be passed 
    an object of this type, which it can handle based on error condition.
    """
    Null = None
    Div0 = None
    Value = None
    Ref = None
    Name = None
    Num = None
    NA = None
    GettingData = None

class _CustomConverter:
    """
    This is the interface class for custom type converters to allow them
    to be called from the Core.
    """
    def __init__(self, callable):
        pass

class _Event:
    def __iadd__(self, handler):
        """
        Registers an event handler function, for example:
            
            event.NewWorkbook += lambda wb_name: print(wb_name)
            
        """
        pass
    def __isub__(self, handler):
        """
        Removes a previously registered event handler function
        """
        pass
    def handlers(self):
        """
        Returns a list of registered handlers for this event
        """
        pass

# Strictly speaking, xloil_core.event is a module but this
# should give the right doc strings
class Event:
    """
    Contains hooks for events driven by user interaction with Excel. The
    events correspond to COM/VBA events and are described in detail at
    `Excel.Application <https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)#events>`_


    Notes:
        * The `CalcCancelled` and `WorkbookAfterClose` event are not part of the 
            Application object, see their individual documentation.
        * Where an event has reference parameter, for example the `cancel` bool in
            `WorkbookBeforeSave`, you need to set the value using `cancel.value=True`.
            This is because python does not support reference parameters for primitive types. 

    Examples
    --------

    ::

        def greet(workbook, worksheet):
            xlo.Range(f"[{workbook}]{worksheet}!A1") = "Hello!"

        xlo.event.WorkbookNewSheet += greet

    """

    AfterCalculate= _Event()
    CalcCancelled= _Event()
    """
    Called when the user interrupts calculation by interacting with Excel.
    """
    NewWorkbook= _Event()
    SheetSelectionChange= _Event()
    SheetBeforeDoubleClick= _Event()
    SheetBeforeRightClick= _Event()
    SheetActivate= _Event()
    SheetDeactivate= _Event()
    SheetCalculate= _Event()
    SheetChange= _Event()
    WorkbookOpen= _Event()
    WorkbookActivate= _Event()
    WorkbookDeactivate= _Event()
    WorkbookAfterClose= _Event()
    """
    Excel's event *WorkbookBeforeClose*, is  cancellable by the user so it is not 
    possible to know if the workbook actually closed.  When xlOil calls 
    `WorkbookAfterClose`, the workbook is certainly closed, but it may be some time
    since that closure happened.

    The event is not called for each workbook when xlOil exits.
    """
    WorkbookBeforeSave= _Event()
    WorkbookBeforePrint= _Event()
    WorkbookNewSheet= _Event()
    WorkbookAddinInstall= _Event()
    WorkbookAddinUninstall= _Event()

event = Event()

class Cache:
    """
    Provides a link to the Python object cache

    Examples
    --------

    ::
        
        @xlo.func
        def myfunc(x):
            return xlo.cache(MyObject(x)) # <- equivalent to .add(...)

        @xlo.func
        def myfunc2(array: xlo.Array(str), i):
            return xlo.cache[array[i]] # <- equivalent to .get(...)

    """

    def add(self, obj):
        """
        Adds an object to the cache and returns a reference string
        based on the currently calculating cell.

        xlOil automatically adds unconvertible returned objects to the cache,
        so this function is useful to force a recognised object, such as an 
        iterable into the cache, or to return a list of cached objects.
        """
        pass

    def get(self, ref:str):
        """
        Fetches an object from the cache given a reference string.
        Returns None if not found
        """
        pass

    def contains(self, ref:str):
        """
        Returns True if the given reference string links to a valid object
        """
        pass

    __contains__ = contains
    __getitem__ = get
    __call__ = add

cache = Cache()

class RtdPublisher:
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

    def __init__(self):
        """
        This __init__ method must be called explicitly by subclasses or 
        pybind will fatally crash Excel.
        """
        pass
    def connect(self, num_subscribers):
        """
        Called by the RtdServer when a sheet function subscribes to this 
        topic. Typically a topic will start up its publisher on the first
        subscriber, i.e. when num_subscribers == 1
        """
        pass
    def disconnect(self, num_subscribers):
        """
        Called by the RtdServer when a sheet function disconnects from this 
        topic. This happens when the function arguments are changed the
        function deleted. Typically a topic will shutdown its publisher 
        when num_subscribers == 0.

        Whilst the topic remains live, it may still receive new connection
        requests, so generally avoid finalising in this method.
        """
        pass
    def stop(self):
        """
        Called by the RtdServer to indicate that a topic should shutdown
        and dependent threads or tasks and finalise resource usage
        """
        pass
    def done(self) -> bool:
        """
        Returns True if the topic can safely be deleted without 
        leaking resources.
        """
        pass
    def topic(self) -> str:
        """
        Returns the name of the topic
        """
        pass

class RtdServer:
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

    def start(self, topic:RtdPublisher):
        """
        Registers an RtdPublisher publisher with this manager. The RtdPublisher receives
        notification when the number of subscribers changes
        """
        pass
    def publish(self, topic:str, value):
        """
        Publishes a new value for the specified topic and updates all subscribers.
        This function can be called even if no RtdPublisher has been started.

        This function does not use any Excel API and is safe to call at any time
        on any thread.
        """
        pass
    def subscribe(self, topic:str):
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
    def peek(self, topic:str, converter=None):
        """
        Looks up a value for a specified topic, but does not subscribe.
        If there is no active publisher for the topic, it returns None.
        If there is no published value, it will return CellError.NA.

        This function does not use any Excel API and is safe to call at
        any time on any thread.
        """
        pass

def register_functions(module, function_holders):
    """
    Register worksheet functions for a specified module. The function_holders
    must be objects created with FuncDescription.create_holder. Generally, there
    is no need to call this directly, it is used by xlOil internals.
    """
    pass

def deregister_functions(module, function_names):
    """
    Deregisters worksheet functions linked to specified module. Generally, there
    is no need to call this directly.
    """
    pass

def get_event_loop():
    """
    Returns the asyncio event loop assoicated with the async background
    worker thread.
    """
    pass

def set_return_converter(conv):
    pass

class CannotConvert(Exception):
    """
    Should be thrown by a return converter when it is unable to handle the 
    provided type.  It does not indicate a fatal condition, as xlOil will
    fallback to another converter, so no message string is required.
    """
    pass

class _CustomReturn:
    def __init__(self, conv):
        pass

class RibbonControl:
    """
    This object is passed to ribbon callback handlers to indicate which control  
    raised the callback.
    """
    @property
    def id(self):
        """
        A string that represents the Id attribute for the control or custom 
        menu item.
        """
        pass
    @property
    def tag(self):
        """
        A string that represents the Tag attribute for the control or custom 
        menu item.
        """
        pass

class RibbonUI:
    """
    Controls an Ribbon and it's associated COM addin
    """
    def connect(self):
        """
        Connects this COM add-in underlying this Ribbon to Excel. Any specified 
        ribbon XML will be passed to Excel.
        """
        pass
    def disconnect(self):
        """
        Unloads the underlying COM add-in and any ribbon customisation.
        """
        pass
    def set_ribbon(xml:str, handlers:dict):
        """
        See `create_ribbon`. This function can only be called when the Ribbon
        is disconnected.
        """
        pass
    def invalidate(id=None):
        """
        Invalidates the specified control: this clears the caches of the
        responses to all callbacks associated with the control. For example,
        this can be used to hide a control by forcing its getVisible callback
        to be invoked.

        If no control ID is specified, all controls are invalidated.
        """
        pass
    def activate(id):
        """
        Activatives the ribbon tab with the specified id.  Returns False if
        there is no Ribbon or the Ribbon is collapsed.
        """
        pass


def create_ribbon(xml:str, handlers:dict) -> RibbonUI:
    """
    Returns a (connected) RibbonUI object which passes the specified ribbon
    customisation XML to Excel.  When the returned object is deleted, it 
    unloads the Ribbon customisation and the associated COM add-in.

    Parameters
    ----------

    xml: str
        A Ribbon XML string, most easily created with a specialised editor.

    handlers: dict
        The ``handlers`` dictionary links callbacks named in the Ribbon XML to
        python functions. Each handler should take a single ``RibbonControl``
        argument which describes the control which raised the callback.

    """
    pass
from ._common import *
import typing

#
# If the xloil_core module can be found, we are being called from an xlOil
# embedded interpreter, so we import the module. Otherwise we define
# skeletons of the imported types to support type-checking, linting,
# auto-completion and documentation.
#
if XLOIL_HAS_CORE:
    import xloil_core         # pylint: disable=import-error
    from xloil_core import (  # pylint: disable=import-error
        CellError, Range, ExcelArray, in_wizard, 
        event, cache, RtdServer, RtdPublisher,
        deregister_functions, get_async_loop,
        ExcelGUI, create_gui, 
        excel_callback, excel_state,
        Caller,
        CannotConvert, 
        from_excel_date,
        insert_cell_image,
        TaskPaneFrame as TaskPaneFrame,
        StatusBar,
        Application, ExcelWindow, Workbook, Worksheet, 
        run, run_async, call, call_async
        )
    if XLOIL_EMBEDDED:
        from xloil_core import (
            workbooks, windows, active_worksheet, active_workbook
            )
else:
    # TODO: how can we synchronise the help here with what you see when you actually import xloil_core

    VT = typing.TypeVar('VT')
    class _Future(typing.Generic[VT]):
        """
        An ``asyncio.Future`` like object which can be awaited and supports
        the `result()` method. This class actually wraps a C++ future so does 
        executes in a separate thread unrelated to an `asyncio` event loop. 
        """
        def __await__(self):
            """
            Returns an iterator which conforms to the async protocol
            """
            ...
        def result(self) -> VT:
            """
            Returns the result of the future or throws the resulting excetion.
            Blocking.
            """
            ...
        def done(self) -> bool:
            """
            Returns True if the future has completed
            """
            ...

    def in_wizard():
        """ 
        Returns true if the function is being invoked from the function wizard: costly functions should 
        exit in this case to maintain UI responsiveness. Checking for the wizard is itself not cheap, so 
        use this sparingly.
        """
        pass

    def excel_callback(func, wait:int=0, retry:int=500, api:str="") -> _Future[object]:
        """
        Schedules a callback to be run in the main thread. Much of the COM API in unavailable
        during the calc cycle, in particular anything which involves writing to the sheet.

        Parameters
        ----------

        func: callable
            A callable which takes no arguments and returns nothing

        retry: int
            Millisecond delay between retries if Excel's COM API is busy, e.g. a dialog box
            is open or it is running a calc cycle.  If zero, does no retry

        wait: int
            Number of milliseconds to wait before first attempting to run this function

        api: str
            Specify 'xll' or 'com' or both to indicate which add-in APIs the call requires.
            The default is 'com' and changing this would only be required in rare cases.
        """
        pass

    class _ExcelState:
        version = int()
        hinstance = int()
        hwnd = int()
        main_thread_id = int()

    def excel_state() -> _ExcelState:
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

    class _AppObject:

        @property
        def name(self) -> str:
            """
            A name for the object. e.g. range address, window caption, workbook name
            """
            ...

        def to_com(lib=None):
            """
            Returns a managed COM object which can be used to invoke Excel's full 
            object model. For details of the available calls see the Microsoft 
            documentation on the Excel Object Model. The ``lib`` used to provide COM
            support can be 'comtypes' or 'win32com'. If omitted, the default is 'comtypes'
            unless changed in the XLL's ini file.
            """
            ...

    class Range(_AppObject):
        """
        Similar to the `Excel.Range <https://docs.microsoft.com/en-us/office/vba/api/excel.range(object)>`_ 
        object, this class allows direct access to an area on a worksheet. It uses similar 
        syntax to Excel's object, supporting the ``cell`` and ``range`` functions, however 
        they are zero-based as per python's standard.

        A Range can be accessed and sliced using the usual syntax (the slice step must be 1):

        ::

            x[1, 1] # The value at (1, 1) as a python type: int, str, float, etc.

            x[1, :] # The second row as another Range object

            x[:-1, :-1] # A sub-range omitting the last row and column

        """
        def range(self, 
                  from_row:int, from_col:int, 
                  num_rows:int=None, num_cols:int=None, 
                  to_row:int=None, to_col:int=None):
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
            pass
        def cell(self, row:int, col:int) -> "Range":
            """ 
            Returns a Range object which consists of a single cell. The indices are zero-based 
            from the top left of the parent range.
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
        @value.setter
        def value(self, value):
            ...
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
            Clears all values and formatting.  Any cell in the range will then have Empty type.
            """
            pass
        def address(self, local=False) -> str:
            """
            Gets the range address in A1 format. The `local` parameter specifies whether
            the workbook and sheet name should be included. For example `local=True` gives
            "[Book1]Sheet1!F37" and `local=False` returns "F37".
            """
            pass
        @property
        def nrows(self) -> int:
            """ Returns the number of rows in the range """
            pass
        @property
        def ncols(self) -> int:
            """ Returns the number of columns in the range """
            pass
        def __getitem__(self, at):
            """ 
            Given a 2-tuple, slices the range to return a sub Range or a single element. Uses
            normal python slicing conventions i.e [left included, right excluded), negative
            numbers are offset from the end.  If the tuple specifies a single cell, returns 
            the value in that cell, otherwise returns a Range object.
            """
            pass
        @property
        def formula(self) -> str:
            """
            Property which returns or sets the cell formula if the range is a single cell 
            or returns the array formula if the entire range contains one. Returns an empty
            string if the range does not contain a formula or array formula.
            """
            ...

        @formula.setter
        def formula(self, value:str):
            ...


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
        an object of this type, which it can handle based on the error condition.
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
        def __init__(self, callable, check_cache=True):
            pass

    class _Event:
        def __iadd__(self, handler):
            """
            Registers an event handler function, for example:
            
            ::

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
        """
        Called after a calculation whether or not it completed or was interrupted
        """
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

        PyBye= _Event()
        """
        Called just before xlOil finalises the python interpreter. All python and xlOil
        functionality is still available. This event is useful to stop threads as it is 
        called before threading module teardown, whereas `atexit` is not.
        """

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

        def add(self, obj, tag:str="", key:str=""):
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
            pass

        def get(self, ref:str, default=None):
            """
            Fetches an object from the cache given a reference string.
            Returns `default` if not found
            """
            pass

        def contains(self, ref:str):
            """
            Returns True if the given reference string links to a valid object
            """
            pass

        def keys(self):
            """
            Returns all cache keys as a list of strings
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

            An Exception object can be passed at the value, this will trigger the debugging
            hook if it is set. The exception string and it's traceback will be published.
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


    def deregister_functions(module, function_names):
        """
        Deregisters worksheet functions linked to specified module. Generally, there
        is no need to call this directly.
        """
        pass

    def get_async_loop():
        """
        Returns the asyncio event loop associated with the async background
        worker thread.  All async / RTD worksheet functions are executed 
        on this event loop.
        """
        pass

    class CannotConvert(Exception):
        """
        Should be thrown by a converter when it is unable to handle the 
        provided type.  In a return converter it may not indicate a fatal 
        condition, as xlOil will fallback to another converter.
        """
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

    class TaskPaneFrame:
        """
            References Excel's base task pane object into which the python GUI can be drawn.
            The methods of this object are safe to call from any thread.  COM must be used on Excel's
            main thread, so the methods all wrap their calls to ensure to this happens. This could lead 
            to deadlocks if the call triggers event  handlers on the main thread, which in turn block 
            waiting for the thread originally calling TaskPaneFrame.
        """
        @property
        def parent_hwnd(self):
            """
            Win32 window handle used to attach a python GUI to a task pane frame
            """
            ...
        @property
        def window(self):
            """
            Gives the window of the document window to which the frame is attached, can be 
            used to uniquely identify the pane
            """
            ...

        @property
        def visible(self) -> bool:
            ...
        @visible.setter
        def visible(self, value: bool):
            ...

        @property
        def size(self) -> typing.Tuple[int, int]:
            """
            Returns the task pane size as a tuple (width, height)
            """
            ...
        @size.setter
        def size(self, value: typing.Tuple[int, int]):
            """
            Sets the task pane size given a tuple (width, height)
            """
            ...

        @property
        def title(self) -> str:
            ...

        def add_event_handler(self, handler):
            ...

    

    class ExcelGUI:
        """
        Controls an Ribbon and it's associated COM addin. The methods of this object are safe
        to call from any thread.  COM must be used on Excel's main thread, so the methods all wrap
        their calls to ensure to this happens. This could lead to deadlocks if the call triggers event 
        handlers on the main thread, which in turn block waiting for the thread originally calling ExcelUI.
        """
        async def connect(self) -> _Future[None]:
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
        async def ribbon(self, xml:str, func_names:dict) -> _Future[None]:
            """
            See ``create_gui``. This function can only be called when the Ribbon
            is disconnected.
            """
            pass
        def invalidate(self, id=None):
            """
            Invalidates the specified control: this clears the caches of the
            responses to all callbacks associated with the control. For example,
            this can be used to hide a control by forcing its getVisible callback
            to be invoked.

            If no control ID is specified, all controls are invalidated.
            """
            pass
        def activate(self, id):
            """
            Activatives the ribbon tab with the specified id.  Returns False if
            there is no Ribbon or the Ribbon is collapsed.
            """
            pass
        def create_task_pane(self, name, creator=None, window=None):
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
            pass

        def task_pane_frame(self, name, window=None, progid=None) -> TaskPaneFrame:
            """
            Used internally to create a custom task pane window which can be populated
            with a python GUI.  Most users should use `create_task_pane(...)` instead.

            A COM `progid` can be specified, but this will prevent using a python GUI
            in the task pane. This is a specialised use case.
            """
            ...

        def com_control(self):
            """
            Returns a pointer to the ActiveX / COM control hosted by the task pane.
            This pointer could be manipulated by win32com or comtypes but this is a 
            specialised use case.
            """
            ...

    async def create_gui(ribbon:str="", func_names:dict=None, name:str=None) -> _Future[ExcelGUI]:
        """
        Returns an ExcelUI object which passes the specified ribbon
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
        pass

    class Caller:
        """
        Captures the caller information for a worksheet function. On construction
        the class queries Excel via the xlfCaller function.
        """
        @property
        def sheet(self):
            """
            Gives the sheet name of the caller or None if not called from a sheet.
            """
            pass
        @property
        def workbook(self):
            """
            Gives the workbook name of the caller or None if not called from a sheet.
            If the workbook has been saved, the name will contain a file extension.
            """
            pass
        def address(self, a1style=False):
            """
            Gives the sheet address either in A1 form: 'Sheet!A1' or RC form: 'Sheet!R1C1'
            """
            pass

    def from_excel_date(value):
        """
        Tries to the convert a given number to a dt.date or dt.datetime assuming it is an 
        Excel date serial number.  Strings are parsed using the current date conversion 
        settings. If dt.datetime is provided, it is simply returned as is.  Raises ValueError
        if conversion is not possible.
        """
        pass

    def insert_cell_image(
        writer, 
        size=None, 
        pos=(0, 0), 
        origin:str=None, 
        compress:bool=True):
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
        pass

    class Worksheet(_AppObject):

        @property
        def parent(self):
            """
            Returns the parent Workbook for this Worksheet
            """
            ...

        def __getitem__(self, at):
            """
            If the argument is a string, returns the range specified by the local address, 
            equivalent to ``at_address``.  
            
            If the argument is a 2-tuple, slices the sheet to return a Range or a single element. 
            Uses normal python slicing conventions, i.e [left included, right excluded), negative
            numbers are offset from the end.
            """
            ...

        def range(self, from_row:int, from_col:int, 
                  num_rows:int=None, num_cols:int=None, 
                  to_row:int=None, to_col:int=None) -> Range:

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
            ...
        def cell(self, row:int, col:int) -> Range:
            """ 
            Returns a Range object which consists of a single cell. The indices are zero-based 
            from the top left of the parent range.
            """
            pass
        def at(self, address:str) -> Range:
            """
            Returns the range specified by the local address, e.g. ``at('B3:D6')``
            """
            ...

        def calculate(self):
            # Calculates this worksheet
            ...

        def activate(self):
            # Makes this worksheet the active sheet
            ...

    class Workbook(_AppObject):

        @property
        def path(self) -> str:
            """
            The full path to the workbook, including the filename
            """
            ...

        def __getitem__(self, sheet_name) -> Worksheet:
            """
            Returns the named worksheet which is part of this workbook (if it exists)
            """
            ...

        def worksheet(self, name) -> Worksheet:
            """
            Returns the named worksheet which is part of this workbook (if it exists)
            """
            ...

        @property
        def worksheets(self) -> list:
            """
            A list of all worksheets which are part of this workbook
            """
            ...

        @property
        def windows(self) -> list:
            """
            The workbook being displayed by this window
            """
            ...

    class ExcelWindow(_AppObject):

        @property
        def hwnd(self) -> int:
            """
            The Win32 API window handle as an integer
            """
            ...
        @property
        def workbook(self) -> Workbook:
            """
            The workbook being displayed by this window
            """
            ...
    
    class _Collection(typing.Generic[VT]):
        """
            An interable collection of workbooks, windows, sheets, etc.
        """
        def __iter__(self):
            ...
        def __getitem__(self, i: str) -> VT:
            ...
        @property
        def active(self) -> VT:
            ...


    workbooks:_Collection[Workbook] = _Collection()
    """
        Collection of all open workbooks as ExcelWorkbook objects.
    
        Examples
        --------

            workbooks['MyBook'].path
            windows.active.workbook.path

    """

    windows:_Collection[ExcelWindow] = _Collection()
    """
        Collection of all open windws as ExcelWindow objects.
    
        Examples
        --------

            workbooks['CaptionMyCaption'].path
            workbooks.active.name

    """

    def active_workbook() -> Workbook:
        """
        Returns the currently active workbook
        """
        ...

    def active_worksheet() -> Worksheet:
        """
        Returns the currently active worksheet
        """
        ...
    
    class StatusBar:
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
        def __init__(self, timeout=0):
            """
            Constructs a StatusBar with a timeout specified in milliseconds.  After the 
            StatusBar context exits, any messages will be cleared after this time
            """
            ...
        def __enter__(self):
            ...
        def __exit__(self, *args):
            ...
        def msg(self, text, timeout=0):
            """
            Posts a status bar message, and if `timeout` is non-zero, clears if after
            the specified number of milliseconds
            """
            ...

    def run(func, *args):
        """
            Calls VBA's `Application.Run` taking the function name and up to 30 arguments.
            This can call any user-defined function or macro but not built-in functions.

            The type and order of arguments expected depends on the function being called.

            Must be called on Excel's main thread, for example in worksheet function or 
            command.
        """
        ...

    def run_async(func, *args) -> _Future[object]:
        """
            Calls VBA's `Application.Run` taking the function name and up to 30 arguments.
            This can call any user-defined function or macro but not built-in functions.

            Calls to the Excel API must be done on Excel's main thread: this async function
            can be called from any thread but will require the main thread to be available
            to return a result.
        """
        ...

    def call(func, *args):
        """
            Calls a built-in worksheet function or command or a user-defined function with the 
            given name. The name is case-insensitive; built-in functions take priority in a name
            clash.
            
            The type and order of arguments expected depends on the function being called.  

            `func` can be built-in function number (as an int) which slightly reduces the lookup overhead

            Must be called from a non-local worksheet function on the main thread.

            *call* can also invoke old-style `macro sheet commands <https://docs.excel-dna.net/assets/excel-c-api-excel-4-macro-reference.pdf`>`_
        """
        ...

    async def call_async(func, *args) -> _Future[object]:
        """
            Calls a built-in worksheet function or command or a user-defined function with the 
            given name.  See ``xloil.call``.

            Calls to the Excel API must be done on Excel's main thread: this async function
            can be called from any thread but will require the main thread to be available
            to return a result.
        """
        ...
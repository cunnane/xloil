import inspect
import functools
import importlib
import typing
import numpy as np
import os
import sys
import traceback

#
# If the xloil_core module can be found, we are being called from an xlOil
# embedded interpreter, so we go ahead and import the module. Otherwise we
# define skeletons of the imported types to support type-checking, linting,
# auto-completion and documentation.
#
if importlib.util.find_spec("xloil_core") is not None:
    import xloil_core         # pylint: disable=import-error
    from xloil_core import (  # pylint: disable=import-error
        CellError, FuncOpts, Range, ExcelArray, in_wizard, log,
        event, cache, RtdServer, RtdPublisher, get_event_loop,
        register_functions, deregister_functions,
        CustomConverter as _CustomConverter) 

else:
    def in_wizard():
        """ 
        Returns true if the function is being invoked from the function 
        wizard. Costly functions should exit in this case. But checking 
        for the wizard is itself not cheap, so use this sparingly.
        """
        pass

    def log(msg, level="info"):
        pass

    class Range:
        """
        Similar to an Excel Range object, this class allows access to an area on a 
        worksheet. It uses similar syntax to Excel's object, supporting the ``cell``
        and ``range`` functions.
        """
        def range(self, from_row, from_col, num_rows=-1, num_cols=-1, to_row=None, to_col=None):
            """ 
            Creates a subrange starting from the specified row and columns and ending
            either at a specified row and column or spannig a number of rows and columns.
            Using negative numbers or num_rows or num_cols means an offset from the end,
            as per the usual python array conventions.
            """
            pass
        def cell(self, row, col):
            """ Returns a Range object which consists of the single cell specified """
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
            Gets the range address in A1 format e.g.
                local=False # [Book1]Sheet1!F37
                local=True  # F37
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
            Given a 2-tuple, slices the range to return a sub Range or a 
            single element.
            """
            pass

    class ExcelArray:
        """
        A view of a internal Excel array which can be manipulated without
        copying the underlying data. It's not a general purpose array class 
        but rather used to create efficiencies in type converters.
        
        It can be accessed and sliced using the usual syntax:
            x[1, 1] # The value at 1,1 as int, str, float, etc.
            x[1, :] # The second row as another ExcelArray
        
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
        form #N/A!, #NAME!, etc passed as a function argument. If your
        function does not use a specific type-converter it may be passed 
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
        `Excel.Appliction <https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)#events>`_


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
        Called when the user interrupts calculation by interacting with Excel.
        """
        CalcCancelled= _Event()
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
        """
        Excel's event *WorkbookBeforeClose*, is  cancellable by the user so it is not 
        possible to know if the workbook actually closed.  When xlOil calls 
        `WorkbookAfterClose`, the workbook is certainly closed, but it may be some time
        since that closure happened.

        The event is not called for each workbook when xlOil exits.
        """
        WorkbookAfterClose= _Event()
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
        An RtdServer sits above an Rtd COM server. Each new RtdServer creates a
        new underlying COM server. The manager connects publishers and subscribers
        for topics, identified by a string. 

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
            """
            pass
        def peek(self, topic:str, converter=None):
            """
            Looks up a value for a specified topic, but does not subscribe.
            If there is no active publisher for the topic, it returns None.
            If there is no published value, it will return CellError.NA.
            """
            pass
    
    def register_functions(module, function_holders):
        pass

    def deregister_functions(module, function_names):
        pass

    def get_event_loop():
        """
        Returns the asyncio event loop assoicated with the async background
        worker thread.
        """
        pass

########################################
# END: XLOIL CORE FORWARD DECLARATIONS #
########################################

"""
Tag used to mark functions to register with Excel. It is added 
by the xloil.func decorator to the target func's __dict__
"""
_META_TAG = "_xloil_func_"
_CONVERTER_TAG = "_xloil_converter_"

"""
This annotation includes all the types which can be passed from xlOil to
a function. There is not need to specify it to xlOil, but it could give 
useful type-checking information to other software which reads annotation.
"""
ExcelValue = typing.Union[bool, int, str, float, np.ndarray, dict, list, CellError]

"""
The special AllowRange annotation allows functions to receive the argument
as an ExcelRange object if appropriate. The argument may still be passed
as another type if it was not created from a sheet reference.
"""
AllowRange = typing.Union[ExcelValue, Range]

class Arg:
    """
    Holds the description of a function argument
    """
    def __init__(self, name, help="", typeof=None, default=None, is_keywords=False):
        self.typeof = typeof
        self.name = str(name)
        self.help = help
        self.default = default
        self.is_keywords = is_keywords

    @property
    def has_default(self):
        """ 
        Since None is a fairly likely default value, this function 
        indicates whether there was a user-specified default
        """
        return self.default is not inspect._empty

def _function_argspec(func):
    """
    Returns a list of Arg for a given function which describe
    the function's arguments
    """
    sig = inspect.signature(func)
    params = sig.parameters
    args = []
    for name, param in params.items():
        if param.kind == param.POSITIONAL_ONLY or param.kind == param.POSITIONAL_OR_KEYWORD:
            spec = Arg(name, default=param.default)
            anno = param.annotation
            if anno is not param.empty:
                spec.typeof = anno
                # Add a little help string based on the type annotation
                if isinstance(anno, type):
                    spec.help = f"({anno.__name__})"
                else:
                    spec.help = f"({str(anno)})"
            args.append(spec)
        elif param.kind == param.VAR_POSITIONAL:
             raise Exception(f"Unhandled argument type positional for {name}")
        elif param.kind == param.VAR_KEYWORD: # can type annotions make any sense here?
            args.append(Arg(name, is_keywords=True))
        else: 
            raise Exception(f"Unhandled argument type for {name}")
    return args


def _get_typeconverter(type_name, from_excel=True):
    # Attempt to find converter with standardised name like `From_int`
    try:
        to_from = 'To' if from_excel else 'From'
        name = f"{to_from}_{type_name}"
        if not hasattr(xloil_core, name):
            name = f"{to_from}_cache"
        return getattr(xloil_core, name)()
        
    except:
        raise Exception(f"No converter {to_from.lower()} {type_name}. Expected {name}")

def converter(typ=typing.Callable, range=False):
    """
    Decorator which declares a function or a class to be a type converter.

    A type converter function is expected to take an argument of type:
    int, bool, float, str, ExcelArray, Range (optional)

    The type converter should return a python object, which could be an 
    ExcelArray or Range.

    A type converter class may take parameters into its constructor
    and hold state.  It should implement __call__ to behave as a type
    converter function.

    Both functions and classes are turned into a class which inherits from 
    ``typ``.  This is to support type hints only.

    If ``range`` is True, the xlOil may pass an ExcelRange or and ExcelArray
    object depending on how the function was invoked.  The type converter should 
    handle both cases consistently.

    Examples
    --------
    
    ::

        @converter(double)
        def arg_sum(x):
            if isinstance(x, ExcelArray):
                return np.sum(x.to_numpy())
            elif isinstance(x, str):
                raise Error('Unsupported')
            return x

        @func
        def pyTest(x: arg_sum):
            return x
            
    """

    def decorate(obj):
        if inspect.isclass(obj):

            class Converter(typ):

                def __init__(self):
                    pass # Keeps linter quiet

                def __call__(self, *args, **kwargs):
                    instance = obj(*args, **kwargs)
                    class Inner(typ or instance.target):
                        _xloil_converter_ = _CustomConverter(instance)
                        allow_range = range
                    # This is a function which returns a class created
                    # by a function-which-takes-a-class and is returned
                    # from another function. Simple!
                    return Inner

            return Converter()

        else:

            class Converter(typ):
                _xloil_converter_ = _CustomConverter(obj)
                allow_range = range

            return Converter

    return decorate


class FuncDescription:
    def __init__(self, func):
        self._func = func
        self.args = _function_argspec(func)
        self.name = func.__name__
        self.help = func.__doc__
        self.is_async = False
        self.rtd = False
        self.macro = False
        self.thread_safe = False
        self.volatile = False
        self.local = None

    def create_holder(self):
        """
        Creates a core object which holds function info, argument converters,
        and a reference to the function object
        """
        
        info = xloil_core.FuncInfo()
        info.args = [xloil_core.FuncArg(x.name, x.help) for x in self.args]
        info.name = self.name
        
        if self.help:
            info.help = self.help
            
        has_kwargs = any(self.args) and self.args[-1].is_keywords

        holder = xloil_core.FuncHolder(info, self._func, has_kwargs)
        
        # Set the arg converters based on the typeof provided for 
        # each argument. If 'typeof' is a xloil typeconverter object
        # it's passed through.  If it is a general python type, we
        # attempt to create a suitable typeconverter
        for i, x in enumerate(self.args):
            if x.is_keywords:
                continue
            # Default option is the generic converter which tries to figure
            # out what to return based on the Excel type
            converter = xloil_core.To_object()
            if x.typeof is not None:
                # If it has this attr, we've already figured out the converter type
                if hasattr(x.typeof, _CONVERTER_TAG):
                    converter = getattr(x.typeof, _CONVERTER_TAG)
                    info.args[i].allow_range = getattr(x.typeof, 'allow_range', False)
                elif x.typeof is AllowRange:
                    info.args[i].allow_range = True
                elif x.typeof is ExcelValue:
                    pass # This the explicit generic type, so do nothing
                elif isinstance(x.typeof, type) and x.typeof is not object:
                    converter = _get_typeconverter(x.typeof.__name__, from_excel=True)
            if x.has_default:
                holder.set_arg_type_defaulted(i, converter, x.default)
            else:
                holder.set_arg_type(i, converter)

        holder.local = True if (self.local is None and not self.is_async) else self.local
        holder.rtd_async = self.rtd and self.is_async

        # TODO: if we are a local func we should reject most FuncOpts
        
        holder.set_opts((FuncOpts.Async if (self.is_async and not self.rtd) else 0) 
                        | (FuncOpts.Macro if self.macro or self.rtd else 0) 
                        | (FuncOpts.ThreadSafe if self.thread_safe else 0)
                        | (FuncOpts.Volatile if self.volatile else 0))

        return holder


def _get_meta(fn):
    return fn.__dict__.get(_META_TAG, None)

import asyncio

def _create_event_loop():
    loop = None
    try:
        loop = asyncio.get_event_loop()
    except RuntimeError:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)

    return loop
        
def async_wrapper(fn):
    """
    Wraps an async function or generator with a function which runs that generator on the thread's
    event loop. The wraped function requires an 'xloil_thread_context' argument which provides a 
    callback object to return a result. xlOil will pass this object automatically to functions 
    declared async.
    """

    @functools.wraps(fn)
    def synchronised(*args, xloil_thread_context, **kwargs):

        loop = get_event_loop()
        cxt = xloil_thread_context

        async def run_async():
            try:
                # TODO: is inspect.isasyncgenfunction expensive?
                if inspect.isasyncgenfunction(fn):
                    async for result in fn(*args, **kwargs):
                        cxt.set_result(result)
                else:
                    result = await fn(*args, **kwargs)
                    cxt.set_result(result)
            except (asyncio.CancelledError, StopAsyncIteration):
                raise
            except Exception as e:
                cxt.set_result(str(e) + ": " + traceback.format_exc())

        cxt.set_task(loop.create_task(run_async()))

    return synchronised    

def _pump_message_loop(control):
    """
    Called internally to run the asyncio message loop. The control object
    allows the loop to be stopped
    """
    loop = get_event_loop()

    async def check_stop():
        while not control.stopped():
            await asyncio.sleep(0.2)
        loop.stop()

    loop.create_task(check_stop())
    loop.run_forever()


def func(fn=None, 
         name=None, 
         help=None, 
         args=None,
         group=None, 
         local=None,
         is_async=False, 
         rtd=False,
         macro=False, 
         thread_safe=False, 
         volatile=False):
    """ 
    Decorator which tells xlOil to register the function in Excel. 
    If arguments are annotated using 'typing' annotations, xlOil will attempt to 
    convert values received from Excel to the specfied type, raising an exception 
    if this is not possible. The currently available types are

    * **int**
    * **float**
    * **str**: Note this disables cache lookup
    * **bool**
    * **numpy arrays**: see Array
    * **CellError**: Excel has various error types such as #NUM!, #N/A!, etc.
    * **None**: if the argument points to an empty cell
    * **cached objects**
    * **datetime.date**
    * **datetime.datetime**
    * **dict / kwargs**: this converter expects a two column array of key/value pairs

    If no annotations are specified, xlOil will pass a type from the first eight above types
    based on the value provided from Excel.

    If a parameter default is given in the function signature, that parameter becomes optional in 
    the declared Excel function.

    Parameters
    ----------

    name: str
        Overrides the funtion name registered with Excel otherwise the function's 
        declared name is used.
    help: str
        Overrides the help shown in the function wizard otherwise the function's 
        doc-string is used.
    args: dict
        A dictionary with key names matching function arguments and values specifying
        information for that argument. The information can be a string, which is 
        interpreted as the help to display in the function wizard or in can be an 
        xloil.Arg object which can contain defaults, help and type information. 
    group: str
        Specifes a category of functions in Excel's function wizard under which
        this function should be placed.
    local: bool
        Functions in a workbook-linked module, e.g. Book1.py, default to 
        workbook-level scope (i.e. not usable outside that workbook) itself. You 
        can override this behaviour with this parameter. It has no effect outside 
        workbook-linked modules.
    macro: bool
        If True, registers the function as Macro Type. This grants the function
        extra priveledges, such as the ability to see un-calced cells and 
        call the full range of Excel.Application functions. Functions which will
        be invoked as Excel macros, i.e. not functions appearing in a cell, should
        be declared with this attribute.
    is_async: bool
        Registers the function as asynchronous. It's better to add the use asyncio's
        'async def' syntax if it is available. Note that only async RTD functions are
        calculated in the background in Excel
    thread_safe: bool
        Declares the function as safe for multi-threaded calculation. The
        function must not make any non-synchronised access to objects outside
        its scope. Since python (at least CPython) is single-threaded there is
        no performance benefit from enabling this.
    volatile: bool
        Tells Excel to recalculate this function on every calc cycle: the same
        behaviour as the NOW() and INDIRECT() built-ins.  Due to the performance 
        hit this brings, it is rare that you will need to use this attribute.

    """

    arguments = locals()
    def decorate(fn):

        _async = is_async
        # If asyncio is not supported e.g. in python 2, this will fail
        # But it doesn't matter since the async wrapper is intended to 
        # hide the async property 
        try:
            if inspect.iscoroutinefunction(fn) or inspect.isasyncgenfunction(fn):
                fn = async_wrapper(fn)
                _async = True
        except NameError:
            pass

        descr = FuncDescription(fn)

        del arguments['fn']
        for arg, val in arguments.items():
            if not arg in ['fn', 'args'] and val is not None:
                descr.__dict__[arg] = val

        if args is not None:
            arg_names = [x.name.casefold() for x in descr.args]
            if type(args) is dict:
                for arg_name, arg_help in args.items():
                    try:
                        i = arg_names.index(arg_name.casefold())
                        descr.args[i].help = arg_help
                    except ValueError:
                        raise Exception(f"No parameter '{arg_name}' in function {fn.__name__}")
            else:
                for arg in args:
                    try:
                        i = arg_names.index(arg.name.casefold())
                        descr.args[i] = arg
                    except ValueError:
                        raise Exception(f"No parameter '{arg_name}' in function {fn.__name__}")
        
        descr.is_async = _async

        fn.__dict__[_META_TAG] = descr
        return fn

    return decorate if fn is None else decorate(fn)

_excel_application_com_obj = None

# TODO: Option to use win32com instead of comtypes?
def app():
    """
        Returns a handle to the Excel.Application object using 
        the comtypes library. The Excel.Application object is the root of
        Excel's COM interface and supports a wide range of operations.
        It is well documented by Microsoft.  Many of the operations are 
        only supported in functions declared as Macro Type.

        Examples
        --------
        To get the name of the active worksheet:
            xlo.app().ActiveSheet.Name

    """
    global _excel_application_com_obj
    if _excel_application_com_obj is None:
        import comtypes.client
        import comtypes
        import ctypes
        clsid = comtypes.GUID.from_progid("Excel.Application")
        obj = ctypes.POINTER(comtypes.IUnknown)(xloil_core.application())
        _excel_application_com_obj = comtypes.client._manage(obj, clsid, None)
    return _excel_application_com_obj
     

class _ArrayType:
    """
    This object can be used in annotations or @xlo.arg decorators
    to tell xlOil to attempt to convert an argument to a numpy array.

    You don't use this type directly, ``Array`` is a static instance of 
    this type, so use the syntax as show in the examples below.

    If you don't specify this annotation, xlOil may still pass an array
    to your function if the user passes a range argument, e.g. A1:B2. In 
    this case you will get a 2-dim Array(object). If you know the data 
    type you want, it is more perfomant to specify it by annotation with 
    ``Array``.

    Examples
    --------

        @xlo.func
        def array1(x: xlo.Array(int)):
            pass

        @xlo.func
        def array2(y: xlo.Array(float, dims=1)):
            pass

        @xlo.func
        def array3(z: xlo.Array(str, trim=False)):
            pass
    
    Methods
    -------

    **(element, dims, trim)** :    
        Element types are converted to numpy dtypes, which means the only supported types are: 
        int, float, bool, str, datetime, object.
        (Numpy has a richer variety of dtypes than this but Excel does not.) 
        
        For the float data type, xlOil will convert #N/As to numpy.nan but other values will 
        causes errors.

    dims : int
        Arrays can be either 1 or 2 dimensional, 2 is the default.  Note the Excel has
        the following behaviour for writing arrays into an array formula range specified
        with Ctrl-Alt-Enter:
        "If you use a horizontal array for the second argument, it is duplicated down to
        fill the entire rectangle. If you use a vertical array, it is duplicated right to 
        fill the entire rectangle. If you use a rectangular array, and it is too small for
        the rectangular range you want to put it in, that range is padded with #N/As."

    trim : bool    
        By default xlOil trims arrays to the last row & column which contain a nonempty
        string or non-#N/A value. This is generally desirable, but can be disabled with 
        this paramter.

    """

    def __call__(self, element=object, dims=2, trim=True):
        name = f"To_Array_{element.__name__}_{dims or 2}d" 
        type_conv = getattr(xloil_core, name)(trim)

        class Arr(np.ndarray):
            _xloil_converter_ = type_conv

        return Arr
        
# Cheat to avoid needing Py 3.7+ for __class_getitem__
Array = _ArrayType() 


try:
    import pandas as pd

    @converter(pd.DataFrame)
    class PDFrame:
        """
        Converter which takes tables with horizontal records to pandas dataframes.

        Examples
        --------

        ::

            @xlo.func
            def array1(x: xlo.PDFrame(int)):
                pass

            @xlo.func
            def array2(y: xlo.PDFrame(float, headings=True)):
                pass

            @xlo.func
            def array3(z: xlo.PDFrame(str, index='Index')):
                pass
    
        Methods
        -------

        **PDFrame(element, headings, index)** : 
            
            element : type
                Pandas performance can be improved by explicitly specifying  
                a type. In particular, creation of a homogenously typed
                Dataframe does not require copying the data.

            headings : bool
                Specifies that the first row should be interpreted as column
                headings

            index : various
                Is used in a call to pandas.DataFrame.set_index()


        """
        def __init__(self, element=None, headings=True, index=None):
            # TODO: use element_type!
            self._element_type = element
            self._headings = headings
            self._index = index

        def __call__(self, x):
            if isinstance(x, ExcelArray):
                df = None
                idx = self._index
                if self._headings:
                    if x.nrows < 2:
                        raise Exception("Expected at least 2 rows")
                    headings = x[0,:].to_numpy(dims=1)
                    data = {headings[i]: x[1:, i].to_numpy(dims=1) for i in range(x.ncols)}
                    if idx is not None and idx in data:
                        index = data.pop(idx)
                        idx = None
                        df = pd.DataFrame(data, index=index)
                    else:
                        # This will do a copy.  The copy can be avoided by monkey
                        # patching pandas - see stackoverflow
                        df = pd.DataFrame(data)
                else:
                    df = pd.DataFrame(x.to_numpy())
                if idx is not None:
                    df.set_index(idx, inplace=True)
                return df
        
            raise Exception(f"Unsupported type: {type(x)!r}")
finally:
    pass


def scan_module(module, workbook_name=None):
    """
        Parses a specified module looking for functions with with the xloil.func 
        decorator and register them. Does not search inside second level imports.

        The argument can be a module object, module name or path string. The module 
        is first imported if it has not already been loaded.
 
        Called by the xlOil C layer to import modules specified in the config.
    """

    if type(module) is str:
        # Not completely sure of the wisdom of adding to sys.path here...
        # But it's difficult to load a module by absolute path
        mod_directory, mod_filename = os.path.split(module)
        if len(mod_directory) > 0 and not mod_directory in sys.path:
            sys.path.append(mod_directory)
        handle = importlib.import_module(os.path.splitext(mod_filename)[0])
    elif (inspect.ismodule(module) and hasattr(module, '__file__')) or module in sys.modules:
        handle = importlib.reload(module)
    else:
        raise Exception(f"scan_module: could not process {str(module)}")

    # Allows 'local' modules to know which workbook they link to
    handle._xl_this_workbook = workbook_name
    
    # Look for functions with an xloil decorator (_META_TAG) and create
    # a function holder object for each of them
    to_register = [_get_meta(x[1]).create_holder() 
                   for x in inspect.getmembers(handle, lambda obj: hasattr(obj, _META_TAG))]

    if any(to_register):
        xloil_core.register_functions(handle, to_register)

    return handle

import inspect
import functools
import importlib
import typing
import numpy as np
import os
import sys

#
# If the xloil_core module can be found, we are being called from an xlOil
# embedded interpreter, so we go ahead and import the module.  Otherwise we
# define skeletons of the imported types to support type-checking, auto- 
# completion and documentation.
#
if importlib.util.find_spec("xloil_core") is not None:
    import xloil_core
    from xloil_core import CellError, FuncOpts, Range, ExcelArray, in_wizard, log
    from xloil_core import CustomConverter as _CustomConverter
    from xloil_core import event, cache

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
        function does not use a specific type-converter it can opt to handle
        these errors.
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

class _ArgSpec:
    """
    Holds the description of a function argument
    """
    def __init__(self, name, default=None, is_keywords = False):
        self.typeof = None
        self.name = str(name)
        self.help = ""
        self.default = default
        self.is_keywords = is_keywords

    @property
    def has_default(self):
        """ 
        Since None is a fairly likely default value this function 
        indicates whether there was a user-specified default
        """
        return self.default is not inspect._empty

def _function_argspec(func):
    """
    Returns a list of _ArgSpec for a given function which describe
    the function's arguments
    """
    sig = inspect.signature(func)
    params = sig.parameters
    args = []
    for name, param in params.items():
        if param.kind == param.POSITIONAL_ONLY or param.kind == param.POSITIONAL_OR_KEYWORD:
            spec = _ArgSpec(name, param.default)
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
            args.append(_ArgSpec(name, is_keywords=True))
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


class _FuncMeta:
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
 
def _create_excelfunc_meta(fn):
    if not hasattr(fn, _META_TAG):
        fn.__dict__[_META_TAG] = _FuncMeta(fn)
    return _get_meta(fn)

def _async_wrapper(fn):
    import asyncio
    """
    Synchronises an 'async def' function. xloil will invoke it
    in a background thread. The C++ layer doesn't doesnt want 
    to have to deal with asyncio event loops.
    """
    @functools.wraps(fn)
    def synchronised(*args, **kwargs):

        # Get current event loop or create one for this thread
        # TODO: because we keep recreating the python threadstate in the C layer
        # all the thread locals get wiped out, so this always creates a loop
        loop = None
        try:
            loop = asyncio.get_event_loop()
        except RuntimeError:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
        
        # Thread context passed from the C++ layer. Remove this 
        # from args intented for the inner function call.
        cxt = kwargs.pop(xloil_core.ASYNC_CONTEXT_TAG)

        task = asyncio.ensure_future(fn(*args, **kwargs), loop=loop)

        while not task.done():
            loop.run_until_complete(asyncio.wait({task}, loop=loop, timeout=1))
            if cxt.cancelled():
                task.cancel()
 
        return task.result()
        
    return synchronised    


def func(fn=None, 
         name=None, 
         help=None, 
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
    group: str
        Specifes a category of functions in Excel's function wizard under which
        this function should be placed.
    local: bool
        For functions in a workbook-associated module, e.g. workbook.py, xlOil
        defaults to scoping their name to the workbook itself. You can override
        this behaviour with this parameter. It has no effect outside associated
        modules.
    macro: bool
        If True, registers the function as Macro Type. This grants the function
        extra priveledges, such as the ability to see un-calced cells and 
        call the full range of Excel.Application functions. Functions which will
        be invoked as Excel macros, i.e. not functions appearing in a cell, should
        be declared with this attribute.
    is_async: bool
        Registers the function as asynchronous. It's better to add the use asyncio's
        'async def' syntax if it is available. Note that async functions aren't
        calculated in the background in Excel: if the user interrupts the calculation
        by interacting with Excel, async functions are cancelled and restarted later.
    thread_safe: bool
        Declares the function as safe for multi-threaded calculation, i.e. the
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
            if inspect.iscoroutinefunction(fn):
                fn = _async_wrapper(fn)
                _async = True
        except:
            pass

        meta = _create_excelfunc_meta(fn)

        del arguments['fn']
        for arg, val in arguments.items():
            if arg is not 'fn' and val is not None:
                meta.__dict__[arg] = val

        meta.is_async = _async

        return fn

    return decorate if fn is None else decorate(fn)


def arg(name, typeof=None, help=None, range=None):
    """ 
        Decorator to specify argument type and help for a function exposed to Excel.
        The help is displayed in the function wizard. 

        This decorator is optional: xloil reads argument types from annotations, 
        and even these are optional. Use this decorator to specify help strings or 
        if 'typing' annotations are unavailable.
    """
    
    def decorate(fn):
        meta = _create_excelfunc_meta(fn) # In case we are called before @func
    
        args = meta.args
    
        # The args are already populated from the function signature 
        # so we're guaranteed at most one match
        try:
            match = next((x for x in args if x.name.casefold() == name.casefold()))
        except:
            raise Exception(f"No parameter '{name}' in function {fn.__name__}")
        
        if typeof is not None: match.typeof = typeof
        if help is not None: match.help = help
        if range is True: match.typeof = AllowRange

        return fn

    return decorate

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

        return Arr;
        
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

    if inspect.ismodule(module):
        handle = importlib.reload(module) 
    elif len(os.path.dirname(module)) > 0:
        # Not completely sure of the wisdom of adding to sys.path here...
        mod_directory = os.path.dirname(module)
        if not mod_directory in sys.path:
            sys.path.append(mod_directory)
        mod_filename = os.path.splitext(os.path.basename(module))[0]
        handle = importlib.import_module(mod_filename)
    else:
        handle = importlib.import_module(module)

    # Allows 'local' modules to know which workbook they link to
    handle._xl_this_workbook = workbook_name
    
    # Look for functions with an xloil decorator (_META_TAG) and create
    # a function holder object for each of them
    to_register = [_get_meta(x[1]).create_holder() 
                   for x in inspect.getmembers(handle, lambda obj: hasattr(obj, _META_TAG))]

    if any(to_register):
        xloil_core.register_functions(handle, to_register)

    return handle

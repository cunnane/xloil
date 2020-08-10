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
        CannotConvert, set_return_converter,
        CustomConverter as _CustomConverter,
        CustomReturn as _CustomReturn)

else:
    from .shadow_core import *
    

"""
Tag used to mark functions to register with Excel. It is added 
by the xloil.func decorator to the target func's __dict__
"""
_META_TAG = "_xloil_func_"

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
    return args, sig.return_annotation


def _get_typeconverter(type_name, from_excel=True):
    # Attempt to find converter with standardised name like `From_int`
    try:
        to_from = 'To' if from_excel else 'From'
        name = f"{to_from}_{type_name}"
        #if not hasattr(xloil_core, name):
        #    name = f"{to_from}_cache"
        return getattr(xloil_core, name)()
        
    except:
        raise Exception(f"No converter {to_from.lower()} {type_name}. Expected {name}")

class _Converter:
     _xloil_converter = None
     _xloil_allow_range = False

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

            #
            # The construction is a little cryptic here to support nice syntax. The
            # target `obj` is intented to be used in a typing expresion like x: obj(...).
            # Hence obj(...) must return something which inherits from the desired type
            # `typ` but is also identifiable to xlOil as a converter.
            #
            class TypingHolder(typ):

                def __init__(self):
                    pass # Keeps linter quiet

                def __call__(self, *args, **kwargs):
                    instance = obj(*args, **kwargs)
                    class Converter(typ or instance.target, _Converter):
                        _xloil_converter = _CustomConverter(instance)
                        _xloil_allow_range = range

                    return Converter

            return TypingHolder()

        else:

            class Converter(typ, _Converter):
                _xloil_converter = _CustomConverter(obj)
                _xloil_allow_range = range

            return Converter

    return decorate

class FuncDescription:
    def __init__(self, func):
        self._func = func
        self.args, self.return_type = _function_argspec(func)
        self.name = func.__name__
        self.help = func.__doc__
        self.is_async = False
        self.rtd = False
        self.macro = False
        self.threaded = False
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
                if inspect.isclass(x.typeof) and issubclass(x.typeof, _Converter):
                    converter = x.typeof._xloil_converter
                    info.args[i].allow_range = x.typeof._xloil_allow_range
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

        if self.return_type is not inspect._empty:
            ret = self.return_type
            if issubclass(ret, _Converter):
                holder.return_converter = _CustomReturn(ret._xloil_converter)
            elif isinstance(x.typeof, type) and x.typeof is not object:
                holder.return_converter = _get_typeconverter(ret.__name__, from_excel=False)
            else:
                pass # TODO: Ignore - warn user?

        holder.local = True if (self.local is None and not self.is_async) else self.local
        holder.rtd_async = self.rtd and self.is_async

        # TODO: if we are a local func we should reject most FuncOpts
        
        holder.set_opts((FuncOpts.Async if (self.is_async and not self.rtd) else 0) 
                        | (FuncOpts.Macro if self.macro or self.rtd else 0) 
                        | (FuncOpts.ThreadSafe if self.threaded else 0)
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
         threaded=False, 
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
        doc-string is used. The wizard cannot display strings longer than 255 chars.
        Longer help string can be retrieved with `xloHelp`
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
        Registers the function as asynchronous. It's better to use asyncio's
        'async def' syntax if it is available. Only async RTD functions are
        calculated in the background in Excel, non-RTD functions will be stopped
        if calculation is interrupted.
    threaded: bool
        Declares the function as safe for multi-threaded calculation. The
        function must be careful when accessing global objects. 
        Since python (at least CPython) is single-threaded there is
        no direct performance benefit from enabling this. However, if you make 
        frequent calls to C-based libraries like numpy or pandas you make
        be able to realise speed gains.
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

        class Arr(np.ndarray, _Converter):
            _xloil_converter = type_conv
            _xloil_allow_range = False

        return Arr
        
# Cheat to avoid needing Py 3.7+ for __class_getitem__
Array = _ArrayType() 


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
        handle = importlib.import_module(mod_filename.replace('.py', ''))
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

class _ReturnConverters:

    _converters = set()

    def add(self, conv):
        self._converters.add(conv)
        xloil_core.set_return_converter(_CustomReturn(self))
    
    def remove(self, conv):
        self._converters.remove(conv)
        if len(self._converters) == 0:
            xloil_core.set_return_converter(None)

    def __call__(self, obj):
        for c in self._converters:
            try:
                return c(obj)
            except (xloil_core.CannotConvert):
                continue
        raise xloil_core.CannotConvert()

return_converters = _ReturnConverters()
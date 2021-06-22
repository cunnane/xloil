import inspect
import functools
import os
import sys
import traceback
from .type_converters import *
from .shadow_core import *


"""
Tag used to mark functions to register with Excel. It is added 
by the xloil.func decorator to the target func's __dict__
"""
_FUNC_META_TAG = "_xloil_func_"
_LANDMARK_TAG = "_xloil_hasfuncs_"

def _insert_landmark(obj):
    module = inspect.getmodule(obj)
    setattr(module, _LANDMARK_TAG, True)

class Arg:
    """
    Holds the description of a function argument. Can be used with the 'func'
    decorator to specify the argument description.
    """
    def __init__(self, name, help="", typeof=None, default=None, is_keywords=False):
        """
        Parameters
        ----------

        name: str
            The name of the argument which appears in Excel's function wizard
        help: str, optional
            Help string to display in the function wizard
        typeof: object, optional
            Selects the type converter used to pass the argument value
        default: object, optional
            A default value to pass if the argument is not specified in Excel
        is_keywords: bool, optional
            Denotes the special kwargs argument. xlOil will expect a two-column array
            in Excel which it will interpret as key, value pairs and convert to a
            dictionary.
        """

        self.typeof = typeof
        self.name = str(name)
        self.help = help
        self.default = default
        self.is_keywords = is_keywords

    @property
    def has_default(self):
        """ 
        Since 'None' is a fairly likely default value, this property indicates 
        whether there was a user-specified default
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


class FuncDescription:
    """
    Used to create the description of a worksheet function to register. 
    External users would not typically use this class directly.
    """
    def __init__(self, func):
        self._func = func
        self.args, self.return_type = _function_argspec(func)
        self.name = func.__name__
        self.help = func.__doc__
        self.is_async = False
        self.rtd = None
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
        for i, arg_info in enumerate(self.args):
            if arg_info.is_keywords:
                continue

            # Determine the internal C++ arg converter to run on the Excel values
            # before they are passed to python.  
            converter = None
            this_arg = info.args[i]
            arg_type = arg_info.typeof

            # If a typing annotation is None or not a type, ignore it.
            # The default option is the generic converter which gives a python 
            # type based on the provided Excel type
            if not isinstance(arg_type, type):
                converter = xloil_core.Read_object()
            else:
                # The ordering of these cases is based on presumed likeliness.
                # First try an internal converter e.g. Read_str, Read_float, etc.
                converter = get_internal_converter(arg_type.__name__)

                # xloil_core.Range is special: the only core class in typing annotations
                if arg_type is Range:
                    this_arg.allow_range = True

                # If internal converter was found, nothing more to do
                if converter is not None:
                    pass
                # A designated xloil @converter type contains the internal converter
                elif is_type_converter(arg_type):
                    converter, this_arg.allow_range = unpack_type_converter(arg_type)
                # ExcelValue is just the explicit generic type, so do nothing
                elif arg_type is ExcelValue:
                    pass 
                elif arg_type is AllowRange:
                    converter = xloil_core.Read_object(), 
                    this_arg.allow_range = True
                # Attempt to find a registered user-converter, otherwise assume the object
                # should be read from the cache 
                else:
                    converter = arg_converters.get_converter(arg_type)
                    if converter is None:
                        converter = xloil_core.Read_Cache()
            log(f"Func '{info.name}', arg '{arg_info.name}' using converter {type(converter)}", level="trace")
            if arg_info.has_default:
                this_arg.optional = True
                holder.set_arg_type_defaulted(i, converter, arg_info.default)
            else:
                holder.set_arg_type(i, converter)

        if self.return_type is not inspect._empty:
            ret_type = self.return_type
            if isinstance(ret_type, type):

                ret_con = None
                if is_type_converter(ret_type):
                    ret_con, _ = unpack_type_converter(ret_type)
                else:
                    ret_con = return_converters.create_returner(ret_type)

                    if ret_con is None:
                        ret_con = get_internal_converter(ret_type.__name__, read_excel_value=False)

                    if ret_con is None:
                        ret_con = Return_object()

                holder.return_converter = ret_con

        # RTD-async is default unless rtd=False was explicitly specified.
        holder.rtd_async = self.is_async and (self.rtd is not False)
        holder.native_async = self.is_async and not holder.rtd_async

        holder.local = True if (self.local is None and not holder.native_async) else self.local

        func_options = ((FuncOpts.Macro if self.macro else 0)
                        | (FuncOpts.ThreadSafe if self.threaded else 0)
                        | (FuncOpts.Volatile if self.volatile else 0))

        if holder.local:
            if func_options != 0:
                log(f"Ignoring func options for local function {self.name}", level='info')
        else:
            holder.set_opts(func_options)
        return holder


def _create_event_loop():
    import asyncio
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    return loop

def _async_wrapper(fn):
    """
    Wraps an async function or generator with a function which runs that generator on the thread's
    event loop. The wrapped function requires an 'xloil_thread_context' argument which provides a 
    callback object to return a result. xlOil will pass this object automatically to functions 
    declared async.

    This function is used by the `func` decorator and generally should not be invoked
    directly.
    """

    import asyncio

    @functools.wraps(fn)
    def synchronised(*args, xloil_thread_context, **kwargs):

        loop = get_event_loop()
        ctx = xloil_thread_context

        async def run_async():
            result = None
            try:
                # TODO: is inspect.isasyncgenfunction expensive?
                if inspect.isasyncgenfunction(fn):
                    async for result in fn(*args, **kwargs):
                        ctx.set_result(result)
                else:
                    result = await fn(*args, **kwargs)
                    ctx.set_result(result)
            except (asyncio.CancelledError, StopAsyncIteration):
                ctx.set_done()
                raise
            except Exception as e:
                ctx.set_result(str(e) + ": " + traceback.format_exc())
                
            ctx.set_done()
            
        ctx.set_task(asyncio.run_coroutine_threadsafe(run_async(), loop))

    return synchronised    

def _pump_message_loop(loop, timeout):
    """
    Called internally to run the asyncio message loop.
    """
    import asyncio

    async def wait():
        await asyncio.sleep(timeout)
    
    loop.run_until_complete(wait())


def func(fn=None, 
         name=None, 
         help=None, 
         args=None,
         group=None, 
         local=None,
         is_async=False, 
         rtd=None,
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

        _is_async = is_async
        if inspect.iscoroutinefunction(fn) or inspect.isasyncgenfunction(fn):
            fn = _async_wrapper(fn)
            _is_async = True

        descr = FuncDescription(fn)

        for arg, val in arguments.items():
            if not arg in ['fn', 'args', 'name', 'help']:
                descr.__dict__[arg] = val
        if name is not None:
            descr.name = name
        if help is not None:
            descr.help = help

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
        
        descr.is_async = _is_async

        # Add the xlOil tags to the function and module
        setattr(fn, _FUNC_META_TAG, descr)
        _insert_landmark(fn)

        return fn

    return decorate if fn is None else decorate(fn)

_excel_application_obj = None


def _get_excel_application_obj():
    import comtypes
    import comtypes.client
    import ctypes

    clsid = comtypes.GUID.from_progid("Excel.Application")
    obj = ctypes.POINTER(comtypes.IUnknown)(xloil_core.application())
    return comtypes.client._manage(obj, clsid, None)

# TODO: Option to use win32com instead of comtypes?
def app():
    """
        Returns a handle to the Excel.Application object using the *comtypes* 
        library. The Excel.Application object is the root of Excel's COM
        interface and supports a wide range of operations. It is well 
        documented by Microsoft, see 
        https://docs.microsoft.com/en-us/visualstudio/vsto/excel-object-model-overview
        and 
        https://docs.microsoft.com/en-us/office/vba/api/excel.application(object).
        
        Many operations using the Application object will only work in 
        functions declared as **macro type**.

        Examples
        --------

        To get the name of the active worksheet:

        ::
            
            @func(macro=True)
            def sheetName():
                return xlo.app().ActiveSheet.Name

    """
    global _excel_application_obj
    if _excel_application_obj is None:
        _excel_application_obj = _get_excel_application_obj()
    return _excel_application_obj
     


class EventsPaused():
    """
    A context manager which stops Excel events from firing whilst
    the context is in scope
    """
    def __enter__(self):
        event.pause()
        return self
    def __exit__(self, type, value, traceback):
        event.allow()



def scan_module(module):
    """
        Parses a specified module to look for functions with with the xloil.func 
        decorator and register them. 
    """

    # We quickly discard modules which do not contain xloil declarations 
    if getattr(module, _LANDMARK_TAG, None) is None:
        return 

    # If events are not paused this function can be entered multiply for the same module
    with EventsPaused() as events_paused:

        if getattr(module, _LANDMARK_TAG) is False:
            return

        log(f"Found xloil functions in {module}", level="debug")

        # Look for functions with an xloil decorator (_META_TAG) and create
        # a function holder object for each of them
        xloil_funcs = inspect.getmembers(module, 
            lambda obj: inspect.isfunction(obj) and hasattr(obj, _FUNC_META_TAG))

        to_register = []
        for f_name, f in xloil_funcs:
            try:
                to_register.append(getattr(f, _FUNC_META_TAG).create_holder())
            except Exception as e:
                log(f"Register failed for {f_name}: {traceback.format_exc()}", level='error')

        if any(to_register):
            xloil_core.register_functions(module, to_register)

        # Unset flag so we don't try to reregister functions
        module.__dict__[_LANDMARK_TAG] = False

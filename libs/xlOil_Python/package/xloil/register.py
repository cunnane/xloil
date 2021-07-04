import inspect
import functools
import os
import sys
import traceback
import importlib.util
from .type_converters import *
from .shadow_core import *
from .com import EventsPaused
from xloil import log

if importlib.util.find_spec("xloil_core") is not None:
    import xloil_core
    from xloil_core import (
        Read_object as _Read_object,
        Read_Cache as _Read_Cache,
        FuncSpec as _FuncSpec
    )
else:
    def _Read_object():
        pass
    def _Read_Cache():
        pass
    class _FuncSpec:
        def __init__(self, *args, **kwargs):
            pass


"""
Tag used to mark modules which contain functions to register. It is added 
by the xloil.func decorator to the module's __dict__ and contains a list
of functions
"""
_LANDMARK_TAG = "_xloil_pending_funcs_"


def _add_pending_fucs(module, objects):
    pending = getattr(module, _LANDMARK_TAG, set())
    pending.update(objects)
    setattr(module, _LANDMARK_TAG, pending)
    
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

    def write_spec(self, this_arg):

        # Set the arg converters based on the typeof provided for 
        # each argument. If 'typeof' is a xloil typeconverter object
        # it's passed through.  If it is a general python type, we
        # attempt to create a suitable typeconverter
        # Determine the internal C++ arg converter to run on the Excel values
        # before they are passed to python.  
        this_arg.name = self.name
        this_arg.help = self.help

        if self.is_keywords:
            return

        arg_type = self.typeof
        converter = 0
        # If a typing annotation is None or not a type, ignore it.
        # The default option is the generic converter which gives a python 
        # type based on the provided Excel type
        if not isinstance(arg_type, type):
            converter = _Read_object()
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
                converter = _Read_object(), 
                this_arg.allow_range = True
            # Attempt to find a registered user-converter, otherwise assume the object
            # should be read from the cache 
            else:
                converter = arg_converters.get_converter(arg_type)
                if converter is None:
                    converter = _Read_Cache()
        if self.has_default:
            this_arg.default = self.default

        assert converter is not None
        this_arg.converter = converter


    @staticmethod
    def override_arglist(arglist, replacements):
        if replacements is None:
            return arglist
        elif not isinstance(replacements, dict):
            replacements = { a.name : a for a in replacements }

        def override_arg(arg):
            override = replacements.get(arg.name, None)
            if override is None:
                return arg
            elif isinstance(override, str):
                arg.help = override
                return arg
            else:
                return override

        return [override_arg(arg) for arg in arglist]

# TODO: Could be replaced by inspect.getfullargspec??
def function_arg_info(func):
    """
    Returns a list of Arg for a given function which describe the function's arguments
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


def find_return_converter(ret_type: type):
    """
    Get an xloil_core return converter for a given type.
    """
    if not isinstance(ret_type, type):
        return None

    ret_con = None
    if is_type_converter(ret_type):
        ret_con, _ = unpack_type_converter(ret_type)
    else:
        ret_con = return_converters.create_returner(ret_type)

        if ret_con is None:
            ret_con = get_internal_converter(ret_type.__name__, read_excel_value=False)

        if ret_con is None:
            ret_con = Return_object()

    return ret_con

def _create_event_loop():
    import asyncio
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    return loop

def async_wrapper(fn):
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


class _WorksheetFunc:
    """
    Decorator class for functions declared using `func` which contains information
    about the Excel function to be registered
    """
    def __init__(self, func, spec):
        self.__wrapped__ = func
        self._xloil_spec = spec
        self.__doc__     = spec.help
        self.__name__    = spec.name
    def __call__(self, *args, **kwargs):
        return self.func(*args, **kwargs)


def func(fn=None, 
         name=None, 
         help="", 
         args=None,
         group="", 
         local=None,
         rtd=None,
         macro=False, 
         threaded=False, 
         volatile=False,
         register=True):
    """ 
    Decorator which tells xlOil to register the function (or callable) in Excel. 
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

    fn: function or Callable:
        Automatically passed when `func` is used as a decorator
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
        be invoked as Excel macros, i.e. not functions called from a cell, should
        be declared with this attribute.
    rtd: bool
        Determines whether a function declared as async uses native or RTD async.
        Only RTD functions are calculated in the background in Excel, native async
        functions will be stopped if calculation is interrupted. Default is True.
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

    def decorate(fn):

        try:
            is_async = False
            if inspect.iscoroutinefunction(fn) or inspect.isasyncgenfunction(fn):
                fn = async_wrapper(fn)
                is_async = True

            func_args, return_type = function_arg_info(fn)

            has_kwargs = any(func_args) and func_args[-1].is_keywords

            # RTD-async is default unless rtd=False was explicitly specified.
            features=""
            if is_async:
                features=("rtd" if rtd is None or rtd else "async")
            elif macro:
                features="macro"
            elif threaded:
                features="threaded"

            # Default to true unless overriden - the parameter is ignored if a workbook
            # has not been linked
            is_local = True if (local is None and not features == "async") else local
            if local and len(features) > 0:
                log(f"Ignoring func options for local function {self.name}", level='info')

            spec = _FuncSpec(
                func = fn,
                nargs = len(func_args),
                name = name if name else "",
                features = features,
                help = help if help else "",
                category = group if group else "",
                volatile = volatile,
                local = is_local,
                has_kwargs = has_kwargs)

            func_args = Arg.override_arglist(func_args, args)

            for i, arg in enumerate(func_args):
                arg.write_spec(spec.args[i])

            if return_type is not inspect._empty:
                spec.return_converter = find_return_converter(return_type)

            log(f"Found func: {str(spec)}", level="debug")

        except Exception as e:
            fn_name = getattr(fn, "__name__", str(fn))
            log(f"Failed determing spec for '{fn_name}': {traceback.format_exc()}", level='error')
         
        if register: # and inspect.isfunction(fn):
            _add_pending_fucs(inspect.getmodule(fn), [spec])

        return _WorksheetFunc(fn, spec)

    return decorate if fn is None else decorate(fn)

   
def _clear_pending_registrations(module):
    """
    Called by the xloil reload hook to start afresh with function registrations
    """
    if hasattr(module, _LANDMARK_TAG):
        delattr(module, _LANDMARK_TAG)

def scan_module(module):
    """
        Parses a specified module to look for functions with with the xloil.func 
        decorator and register them. Rather than call this manually, it is easer
        to import xloil.importer which registers a hook on the import function.
    """

    # We quickly discard modules which do not contain xloil declarations 
    pending_funcs = getattr(module, _LANDMARK_TAG, None) 
    if pending_funcs is None or not any(pending_funcs):
        return 

    # If events are not paused this function can be entered multiply for the same module
    with EventsPaused() as events_paused:

        log(f"Found xloil functions in {module}", level="debug")

        xloil_core.register_functions(list(pending_funcs), module, append=False)
                                      
        pending_funcs.clear()
        
def register_functions(funcs, module=None, append=True):
    """
        Registers the provided callables and associates tehm with the given modeule

        Parameters
        ----------

        funcs: iterable
            An iterable of `_WorksheetFunc` (a callable decorated with `func`), callables or
            `_FuncSpec`.  A callable is registered by using `func` with the default settings.
            Passing one of the other two allows control over the registration such as changing
            the function or argument names.
        module: python module
            A python module which is the source of the functions. If this module is edited 
            it is automatically reloaded and the functions re-registered. Passing None disables
            this behaviour
        append: bool
            Whether to append to or overwrite any existing functions associated with the module
    """

    # Check if we have a _FuncSpec, else call the decorator to get one 
    def to_spec(f):
        if isinstance(f, _FuncSpec):
            return f
        elif isinstance(f, _WorksheetFunc):
            return f._xloil_spec
        else:
            return func(f, register=False)._xloil_spec

    to_register = [to_spec(f) for f in funcs]

    # We don't know if the module is in the process of loading. Since scan_module will
    # overwrite all existing functions, we both register now and add to the pending list 
    # Registering the same function twice is optimised by xlOil to avoid overhead
    # TODO: check we are called from exec_module for the matching module object
    _add_pending_fucs(module, to_register)
    xloil_core.register_functions(to_register, module, append)

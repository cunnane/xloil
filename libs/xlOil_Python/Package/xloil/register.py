import inspect
import functools
from .type_converters import *
from ._core import *
from .logging import *
from .func_inspect import Arg
import contextvars
import typing

from xloil_core import (
    _Read_object,
    _Read_Cache,
    _Return_Cache,
    _FuncSpec,
    _FuncArg,
    _register_functions
)

_LANDMARK_TAG = "_xloil_pending_funcs_"

# typing.List (no square brackets) is a _SpecialGenericAlias in Py 3.9, but this type 
# doesn't exist in 3.8
_TYPING_GENERIC_CLASSES = (typing._SpecialGenericAlias, typing._GenericAlias) \
    if hasattr(typing, "_SpecialGenericAlias") else typing._GenericAlias

"""
    Tag used to mark modules which contain functions to register. It is added 
    by the xloil.func decorator to the module's __dict__ and contains a list
    of functions
"""

def _add_pending_funcs(module, objects):
    pending = getattr(module, _LANDMARK_TAG, set())
    pending.update(objects)
    setattr(module, _LANDMARK_TAG, pending)
 
def arg_to_funcarg(arg: Arg) -> _FuncArg:

    # Set the arg converters based on the typeof provided for 
    # each argument. If 'typeof' is a xloil typeconverter object
    # it's passed through.  If it is a general python type, we
    # attempt to create a suitable typeconverter
    # Determine the internal C++ arg converter to run on the Excel values
    # before they are passed to python.

    flags = ""
    arg_type = arg.typeof
    converter = None
    allow_range = False

    if arg.kind == Arg.KEYWORD_ARGS:
        flags = "keywords"
        return _FuncArg(arg.name, arg.help, None, "keywords")

    # FastArray is handled separately as converters/defaults are not used 
    if arg_type is FastArray:
        flags = "array"
        if arg.has_default:
            log.warn(f"Defaults not supported for FastArray parameter {arg.name}")
        return _FuncArg(arg.name, arg.help, None, "array")

    if arg_type is ExcelValue:
        # ExcelValue is just the explicit generic type
        converter = _Read_object()
    
    elif arg_type is AllowRange:
        converter = _Read_object()
        allow_range = True

    # Prior to Py 3.9, type annotions for builtins looked like typing.List[int]
    # We resolve this to the base type. 
    # TODO: user-defined converters for these types won't get called
    elif isinstance(arg_type, _TYPING_GENERIC_CLASSES):
        type_name = arg_type._name.lower()
        converter = get_converter(type_name)

    elif not isinstance(arg_type, type):
        converter = _Read_object()

    else:
        converter = get_converter(arg_type.__name__)

        # xloil_core.Range is special: the only core class in typing annotations
        if arg_type is Range:
            allow_range = True
            
        # If internal converter was found, nothing more to do
        if converter is not None:
            pass
        # A designated xloil @converter type contains the internal converter
        elif unpack_arg_converter(arg_type) is not None:
            converter, allow_range = unpack_arg_converter(arg_type)
        # Attempt to find a registered user-converter, otherwise assume the object
        # should be read from the cache 
        else:
            converter = arg_converters.get_converter(arg_type)
            if converter is None:
                converter = _Read_Cache()

    if allow_range:
        flags = "range"
 
    if arg.kind == Arg.VARIABLE_ARGS:
        flags += ",vargs"

    result = _FuncArg(arg.name, arg.help, converter, flags)
    if arg.has_default:
        result.default = arg.default

    log.debug(f"Interpreted arg '{arg}' => '{result}'")

    return result


def find_return_converter(ret_type: type):
    """
    Get an xloil_core return converter for a given type.
    """
    if not isinstance(ret_type, type):
        return None

    ret_con = unpack_return_converter(ret_type)
    if ret_con is None:
        # TODO: can we chain this with 'or' maybe?
        ret_con = return_converters.create_returner(ret_type)

        if ret_con is None:
            ret_con = get_converter(ret_type.__name__, read=False)

        if ret_con is None:
            ret_con = Return_object()

    return ret_con

def _logged_wrapper(func):
    """
    Wraps func so that any errors are logged. Invoked from the core.
    """
    def logged_func(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            log_except(f"Error during {func.__name__}")
    return logged_func

async def _logged_wrapper_async(coro):
    """
    Wraps coroutine so that any errors are logged. Invoked from the core.
    """
    try:
        return await coro
    except Exception as e:
        log_except(f"Error during coroutine")

# This is a thread-local variable to get Caller to behave like a static
# but work properly on different threads and when used in an async funcion
# where normally xlfCaller is not available.
_async_caller = contextvars.ContextVar('async_caller')

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

    def __new__(self, *args, **kwargs):
        global _async_caller
        override = _async_caller.get(None)
        return override or xloil_core.Caller(*args, **kwargs)
    

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
    import traceback
    @functools.wraps(fn)
    def synchronised(xloil_thread_context, *args, **kwargs):

        ctx = xloil_thread_context
        is_async_generator = inspect.isasyncgenfunction(fn)
        
        async def run_async():
            _async_caller.set(ctx.caller)
            try:
                if is_async_generator:
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
            
        ctx.set_task(asyncio.run_coroutine_threadsafe(run_async(), ctx.loop))

    return synchronised

class _WorksheetFunc:
    """
    Decorator class for functions declared using `func`. The class contains 
    the descriptions of the Excel function to be registered
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
         macro=True, 
         command=False,
         threaded=False,
         volatile=False,
         is_async=False,
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
    * **dict / kwargs**: this converter expects a two column array of key/value
      pairs
    * **arg list / *args**

    If no annotations are specified, xlOil will pass a type from the first 
    eight above types based on the value provided from Excel.

    If a parameter default is given in the function signature, that parameter 
    becomes optional in the declared Excel function.

    If keyword args (`**kwargs`) are specified, xlOil expects a two-column 
    array of (string, value) to be passed. For variable args (`*args`) xlOil
    adds a large number of trailing optional arguments. The variable argument 
    list is ended by the first missing argument.  If both *kwargs* and *args* 
    are specified, their order is reversed in the Excel function declaration.

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
        interpreted as the help to display in the function wizard or it can be an
        `xloil.Arg` object which can contain defaults, help and type information. 
    group: str
        Specifes a category of functions in Excel's function wizard under which
        this function should be placed.
    local: bool
        Functions in a workbook-linked module, e.g. Book1.py, default to 
        workbook-level scope (i.e. not usable outside that workbook) itself. You 
        can override this behaviour with this parameter. It has no effect outside 
        workbook-linked modules.
    macro: bool
        If True, registers the function as *Macro Sheet Type*. This grants the 
        function extra priveledges, such as the ability to see un-calced cells  
        and call Excel.Application functions. This is not the same as a 'VBA Macro' 
        which is a *command*.  *Threaded* functions cannot be declared as macro type.
    command: bool
        If True, registers this as a *Command*.  Commands are run outside the 
        calculation cycle on Excel's main thread, have full permissions on the Excel
        object model and do not return a value. They correspond to VBA's 'Macros' or
        'Subroutines'.  Unless declared *local*,  XLL commands are hidden and 
        not displayed in dialog boxes for running macros, although their names 
        can be entered anywhere a valid command name is required.  Commands cannot
        currently be run async since their primary use is writing to the workbook
        which requires running on Excel's main thread.
    rtd: bool
        Determines whether a function declared as async uses native or RTD async.
        Only RTD functions are calculated in the background in Excel, native async
        functions will be stopped if calculation is interrupted. Default is True.
    threaded: bool
        Declares the function as safe for multi-threaded calculation. The
        function must be careful when accessing global objects.  Since python
        (at least CPython) is single-threaded there is no direct performance
        benefit from enabling this. However, if you make frequent calls to
        C-based libraries like numpy or pandas you make be able to realise
        speed gains.
    volatile: bool
        Tells Excel to recalculate this function on every calc cycle: the same
        behaviour as the NOW() and INDIRECT() built-ins.  Due to the performance 
        hit this brings, it is rare that you will need to use this attribute.
    is_async: bool
        If true, manually creates an async function. This means your function
        must take a thread context as its first argument and start its own async
        task similar to ``xloil.async_wrapper``.  Generally this parameter should
        not be used and async functions declared using the normal `async def` syntax.
    """

    def decorate(fn):

        try:
            func_args, return_type = Arg.full_argspec(fn)

            # For async/rtd functions, the target function passed to the FuncSpec 
            # should not be the 'fn' argument to this decorator. We also wrap threaded
            # functions if the debugpy debugger is present
            target_func = fn
            is_coroutine = inspect.iscoroutinefunction(fn)
            is_asyncgen = inspect.isasyncgenfunction(fn)

            # Coroutines need to be "synchronised" using an async/rtd return object.
            # Manual async functions need their arguments shuffled as the first one
            # is the async return.
            if is_coroutine or is_asyncgen:
                target_func = async_wrapper(target_func)
            elif is_async:
                func_args = func_args[1:]
            elif threaded:
                from .debug import _debugpy_activate_thread_decorator
                target_func = _debugpy_activate_thread_decorator(fn)

            # Determine the 'features' string based on our bool flags
            features = []
            local_allowed = True

            if threaded:
                features.append("threaded")
                local_allowed = False

            if (is_async or is_coroutine or is_asyncgen):
                # RTD-async is default unless rtd=False was explicitly specified.
                if rtd == False:
                    features.append("async")
                    local_allowed = False
                else:
                    features.append("rtd")

            if command: 
                features.append("command")

            if return_type is FastArray:
                if any(features):
                    raise ValueError(f"FastArray not compatible with {','.join(features)}")
                features.append('fastarray')
                
            elif macro and not any(features):
                features.append("macro")

            if local == True and not local_allowed:
                raise ValueError(f"'threaded' or 'async' functions cannot be 'local'")

            # is_local defaults to true unless overridden - the parameter is ignored if 
            # a workbook is not linked to the declaring module
            is_local = True if local is None else local

            # Optional overrides of function arg information that we read by reflection

            func_args = Arg.override_arglist(func_args, args)
            core_argspec = [arg_to_funcarg(arg) for arg in func_args]

            spec = _FuncSpec(
                func = target_func,
                args = core_argspec,
                name = name if name else fn.__name__,
                features = ','.join(features),
                help = help if help else (fn.__doc__ if fn.__doc__ else ""),
                category = group if group else "",
                volatile = volatile,
                local = is_local)

            if return_type is not None and return_type is not FastArray:
                spec.return_converter = find_return_converter(return_type)

            # If we are decorating a class, we assume we want to call the class
            # ctor and return the new object as a cache reference, rather than 
            # trying to convert to Excel values
            elif inspect.isclass(fn):
                spec.return_converter = _Return_Cache()

            log(f"Declared excel func: {str(spec)}", level="debug")
  
            if register: # and inspect.isfunction(fn):
                _add_pending_funcs(inspect.getmodule(fn), [spec])

            return _WorksheetFunc(fn, spec)

        except Exception as e:
            fn_name = getattr(fn, "__name__", str(fn))
            log_except(f"Failed determing spec for '{fn_name}'")
            return fn

    return decorate if fn is None else decorate(fn)

   
def _clear_pending_registrations(module):
    """
    Called by the xloil reload hook to start afresh with function registrations
    """
    if hasattr(module, _LANDMARK_TAG):
        delattr(module, _LANDMARK_TAG)


import threading
_scan_module_mutex = threading.Lock()

def scan_module(module, addin=None):
    """
        Parses a specified module to look for functions with with the 
        `xloil.func` decorator and register them. Rather than call this
        manually, it is easer to import xloil.importer which registers a 
        hook on the import function.
    """

    # Enabling this generates a large amount of unhelpful output
    # log.trace(f"Scanning module {module}")

    # We quickly discard modules which do not contain xloil declarations 
    pending_funcs = getattr(module, _LANDMARK_TAG, None) 
    if pending_funcs is None or not any(pending_funcs):
        return 0

    with _scan_module_mutex:
        # Check the pending funcs haven't been processed by another thread then copy 
        # and clear. Other threads can enter this function triggered by Excel's events
        if not any(pending_funcs):
            return 0
        func_list = list(pending_funcs)
        pending_funcs.clear()

        log.debug("Found %d xloil functions in %s", len(func_list), module)

        if addin is None:
            from .importer import source_addin
            addin = source_addin()

        _register_functions(func_list, module, addin, append=False)

        return len(func_list)


def register_functions(funcs, module=None, append=True):
    """
        Registers the provided callables and associates them with the given 
        module

        Parameters
        ----------

        funcs: iterable
            An iterable of `_WorksheetFunc` (a callable decorated with `func`), 
            callables or `_FuncSpec`.  A callable is registered by using `func`
            with the default settings. Passing one of the other two allows control 
            over the registration such as changing the function or argument names.
        module: python module
            A python module which contains the source of the functions (it does 
            not have to equal the module calling this function). If this module is 
            edited it is automatically reloaded and the functions re-registered.
            Passing None disables this behaviour.
        append: bool
            Whether to append to or overwrite any existing functions associated 
            with the module
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
    if module:
        _add_pending_funcs(module, to_register)

    from .importer import source_addin
    addin = source_addin()

    _register_functions(to_register, module, addin, append)


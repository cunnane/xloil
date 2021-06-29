import inspect
import functools
import os
import sys
import traceback
from .type_converters import *
from .shadow_core import *


"""
Tag used to mark functions to register with Excel. It is added by the
xloil.func decorator to the target func's __dict__. It contains a FuncSpec
"""
_FUNC_META_TAG = "_xloil_func_"

"""
Tag used to mark modules which contain functions to register. It is added 
by the xloil.func decorator to the module's __dict__ and contains a list
of functions
"""
_LANDMARK_TAG = "_xloil_pending_funcs_"


def _insert_landmark(obj):
    module = inspect.getmodule(obj)
    pending = getattr(module, _LANDMARK_TAG, None)
    if pending:
        pending.append(obj)
    else:
        setattr(module, _LANDMARK_TAG, [obj])
    
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

        import xloil_core

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
        if self.has_default:
            this_arg.default = self.default

        log(f"Got here: {self.name}, {str(converter)}")
        #assert converter is not None
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


def func(fn=None, 
         name=None, 
         help="", 
         args=None,
         group="", 
         local=None,
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

            spec = xloil_core.FuncSpec(
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

            # Add the xlOil tags to the function and module
            setattr(fn, _FUNC_META_TAG, spec)
            _insert_landmark(fn)

        except Exception as e:
            log(f"Failed determing spec for '{fn.__name__}': {traceback.format_exc()}", level='error')
            
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

    pending_funcs = getattr(module, _LANDMARK_TAG, None) 
    if pending_funcs is None or not any(pending_funcs):
        return 

    # If events are not paused this function can be entered multiply for the same module
    with EventsPaused() as events_paused:

        log(f"Found xloil functions in {module}", level="debug")

        to_register = [getattr(f, _FUNC_META_TAG) for f in pending_funcs]
        
        # Unset flag so we don't try to reregister functions
        setattr(module, _LANDMARK_TAG, [])

        xloil_core.register_functions(module, to_register)


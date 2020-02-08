
import inspect #from standard library
import functools
import importlib
import xloil_core
from xloil_core import CellError, FuncOpts
import numpy as np
import os
import sys

class TypeConvertor:
    pass

"""
    Magic tag which we use to find functions to register with Excel
    It is added by the xloil.func decorator to the target func's 
    __dict__
"""
_META_TAG = "_xlOilFunc_"


class ArgSpec:
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
    Returns a list of ArgSpec for a given function which describe
    the function's arguments
    """
    sig = inspect.signature(func)
    params = sig.parameters
    args = []
    for name, param in params.items():
        if param.kind == param.POSITIONAL_ONLY or param.kind == param.POSITIONAL_OR_KEYWORD:
            spec = ArgSpec(name, param.default)
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
            args.append(ArgSpec(name, is_keywords=True))
        else: 
            raise Exception(f"Unhandled argument type for {name}")
    return args


def _get_typeconverter(type_name, from_excel=True):
    # Attempt to find converter with standardised name
    try:
        to_from = 'from' if from_excel else 'to'
        name = f"{type_name}_{to_from}_Excel"
        if not hasattr(xloil_core, name):
            name = f"cached_{to_from}_Excel"
        return getattr(xloil_core, name)()
        
    except:
        raise Exception(f"No converter for {type_name} {to_from} Excel. Expected {name}")


class FuncMeta:
    def __init__(self, func):
        self._func = func
        self.args = _function_argspec(func)
        self.name = func.__name__
        self.help = func.__doc__
        self.is_async = False
        self.macro = False
        self.thread_safe = False
        self.volatile = False
        
    def create_holder(self):
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
            converter = xloil_core.object_from_Excel()
            if x.typeof is not None:
                # If it has this attr, we've already figured out the converter type
                if hasattr(x.typeof, "_xloil_type_info"):
                    converter = _get_typeconverter(x.typeof._xloil_type_info, from_excel=True)
                elif isinstance(x.typeof, type) and x.typeof is not object:
                    converter = _get_typeconverter(x.typeof.__name__, from_excel=True)
            if x.has_default:
                holder.set_arg_type_defaulted(i, converter, x.default)
            else:
                holder.set_arg_type(i, converter)

        holder.set_opts((FuncOpts.Async if self.is_async else 0) 
                        | (FuncOpts.Macro if self.macro else 0) 
                        | (FuncOpts.ThreadSafe if self.thread_safe else 0)
                        | (FuncOpts.Volatile if self.volatile else 0))

        return holder


class _xloilArray(np.ndarray):
    """
        Should never be invoked directly. It exists to ensure Array[...] can return a type.
        This allows intellisense to work when it is used as an annotations and is 
        consistent with the 'typing' module
    """
    def __init__(self, *args, **kwargs):
        # TODO: following doesn't work for some reason involving numpy & C-extensions
        #super().__init__(*args, **kwargs)
        self._pytype = kwargs['dtype']
        self._set_array(self._pytype)
    
    def __call__(self, dims=None, trim=True):
        return self._set_array(self._pytype, dims, trim)

    def __str__(self):
        return f"Array[{self.dtype}]"

    def _set_array(self, elem_type=object, dims=None, trim=True):
        self.shape = (1,) if dims == 1 else (1,1)
        self.dtype = elem_type
        self._xloil_type_info = f"Array_{elem_type.__name__}_{dims or 2}d"
        return self

class ArrayType:
    def __getitem__(self, elem_type=object):
        return _xloilArray(dtype=elem_type, shape=(1,1))


# Cheat to avoid needing Py 3.7+ for __class_getitem___        
Array = ArrayType()

def _get_meta(fn):
    return fn.__dict__.get(_META_TAG, None)
 
def _create_excelfunc_meta(fn):
    if not hasattr(fn, _META_TAG):
        fn.__dict__[_META_TAG] = FuncMeta(fn)
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
        loop = None
        try:
            loop = asyncio.get_event_loop()
        except RuntimeError:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
        
        # Thread context passed from the C++ layer. Remove this 
        # from args intented for the inner function call.
        cxt = kwargs.pop("xloil_thread_context")

        task = asyncio.ensure_future(fn(*args, **kwargs), loop=loop)

        while not task.done():
            loop.run_until_complete(asyncio.wait({task}, loop=loop, timeout=1))
            if cxt.cancelled():
                task.cancel()
 
        return task.result()
        
    return synchronised    


def func(fn=None, name=None, help=None, group=None, is_async=False, macro=False, thread_safe=False, volatile=False):
    """ 
    Decorator which tells xlOil to register the function in Excel.
        *name*: overrides the funtion name registered with Excel otherwise
                the function's declared name is used.
        *help*: overrides the help shown in the function wizard otherwise
                the function's doc-string is used.
        *group*: specifes a category of functions in Excel's function wizard.
        *is_async*: registers the function as async. It's better to add the 
                    'async' keyword to the function declaration if possible.
    """

    arguments = locals()
    def decorate(fn):

        _async = is_async
        # If asyncio is not supported e.g. python 2, this will fail
        # But it doesn't matter since the async wrapper is intended to 
        # removes the async property
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

def arg(name, typeof=None, help=None):
    """ 
    Decorator to specify argument type and help for a function exposed to Excel 
    If arg is not called for a given argument its type will be inferred from
    annotations or be of a generic type corresponding to the argument passed from
    Excel.
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
    
        return fn

    return decorate

_excel_application_com_obj = None

# TODO: Option to use win32com instead of comtypes?
def app():
    global _excel_application_com_obj
    if _excel_application_com_obj is None:
        import comtypes.client
        import comtypes
        import ctypes
        clsid = comtypes.GUID.from_progid("Excel.Application")
        obj = ctypes.POINTER(comtypes.IUnknown)(xloil_core.application())
        _excel_application_com_obj = comtypes.client._manage(obj, clsid, None)
    return _excel_application_com_obj
     

def _import_from_path(path, module_name=None):
    import importlib.util
    if module_name is None:
        module_name = "xloil." + os.path.splitext(os.path.basename(path))[0]
    # This recipe is copied from the importlib documentation
    spec = importlib.util.spec_from_file_location(module_name, path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = mod
    spec.loader.exec_module(mod)
    return mod


def scan_module(m):
    """
        If objects are not modules, treats them as names as attempt to import them 
        Then looks for functions in the specified modules with the xloil.func meta 
        and registers them. Does not search inside second level imports.

        Called by the xlOil C layer to import  modules specified in the config.
    """

    if inspect.ismodule(m):
        # imp.reload throws strange errors if the module in not on sys.path
        # e.g. https://stackoverflow.com/questions/27272574/imp-reload-nonetype-object-has-no-attribute-name
        # So handle modules we loaded with absolute path separately
        if m.__name__.startswith("xloil"):
            handle = _import_from_path(m.__spec__.origin, m.__name__)
        else:
            handle = importlib.reload(m) 
    elif len(os.path.dirname(m)) > 0:
        handle = _import_from_path(m)
    else:
        handle = importlib.import_module(m)

    to_register = [_get_meta(x[1]).create_holder() for x in inspect.getmembers(handle, lambda obj: hasattr(obj, _META_TAG))]
    if any(to_register):
        xloil_core.register_functions(handle, to_register)

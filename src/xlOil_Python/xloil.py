
"""
xlOil Python
============

The Python plugin for xlOil allows creation of Excel functions and macros backed by Python
code.

To use the plugin, you usually need an entry in the xloil.ini file to set Python paths.

The plugin can load a specified list of module names, adding functions to Excel's global
name scope, like an Excel addin.  The plugin can also look for modules of the form
<workbook_name>.py and load these too.  Any module which contains Excel functions is 
watched for file modifications so code changes are reflected immediately in Excel.

Have a look at `<root>/test/PythonTest.py` for lots of examples. 

Concepts: Cached Objects
------------------------
If xlOil cannot convert a returned python object to Excel, it will place it in an object
cache and return a cache reference string of the form
``UniqueChar[WorkbookName]SheetName!CellRef,#``
If a string of this kind if encountered when reading function arguments, xlOil tries to 
fetch the corresponding python object. With this mechanism you can pass python objects 
opaquely between functions. 

xlOil core also implements a cache for Excel values, which is mostly useful for passing 
arrays. The function ``=xloRef(A1:B2)`` returns a cache string similar to the one used
for Python objects. These strings are automatically looked up when parsing function 
arguments.

"""
import inspect
import functools
import importlib
import typing
import numpy as np
import os
import sys

try:
    import xloil_core
    from xloil_core import CellError, FuncOpts, Range, in_wizard, log
except Exception:
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
        Similar to an Excel Range object, this class allows access to
        an area on a worksheet.
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
            pass
        def clear(self):
            """
            Sets all values in the range to the Nil/Empty type
            """
            pass
        def address(self,local=False):
            """
            Gets the range address in A1 format e.g.
            local=False: [Book1]Sheet1!F37
            local=True:  F37
            """
            pass
        @property
        def num_rows(self):
            """ Returns the number of rows in the range """
            pass
        @property
        def num_cols(self):
            """ Returns the number of columns in the range """
            pass

    class CellError:
        """
        Enum-type class created when an Excel error condition of the 
        form #FOO! is passed a a function argument.
        """
        Null = None
        Div0 = None
        Value = None
        Ref = None
        Name = None
        Num = None
        NA = None
        GettingData = None




"""
Tag used to mark functions to register with Excel. It is added 
by the xloil.func decorator to the target func's __dict__
"""
_META_TAG = "_xlOilFunc_"

ExcelValue = typing.Union[int, str, float, np.ndarray, dict, list]
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
    # Attempt to find converter with standardised name
    try:
        to_from = 'from' if from_excel else 'to'
        name = f"{type_name}_{to_from}_Excel"
        if not hasattr(xloil_core, name):
            name = f"cached_{to_from}_Excel"
        return getattr(xloil_core, name)()
        
    except:
        raise Exception(f"No converter for {type_name} {to_from} Excel. Expected {name}")


class _FuncMeta:
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
                elif x.typeof is AllowRange:
                    info.args[i].allow_range = True
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


def func(fn=None, 
         name=None, 
         help=None, 
         group=None, 
         is_async=False, 
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
    macro: bool
        If True, registers the function as Macro Type. This grants the function
        extra priveledges, such as the ability to see un-calced cells and 
        call the full range of Excel.Application functions. Functions which will
        be invoked as Excel macros, i.e. not functions appearing in a cell, should
        be declared with this attribute
    is_async: bool
        Registers the function as asynchronous. It's better to add the use asyncio's
        'async def' syntax if it is available. Note that async functions aren't
        calculated in the background in Excel: if the user interrupts the calculation
        by interacting with Excel, async functions are cancelled and restarted later.
    thread_safe: bool
        Declares the function as safe for multi-threaded calculation, i.e. the
        function must not make any non-synchronised access to objects outside
        its scope. Since python (at least CPython) is single-threaded there is
        no performance benefit from enabling this
    volatile: bool
        Tells Excel to recalculate this function on every calc cycle: the same
        behaviour as the NOW() and INDIRECT() built-ins.  Due to the performance 
        hit this brings, it is rare that you will need to use this attribute.

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
     

## TODO: implement
def add_cache():
    pass

## TODO: implement
def fetch_cache():
    pass

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
    """
    This object can be used in annotations or @xlo.arg decorators
    to tell xlOil to attempt to convert an argument to a numpy array.

    You don't use this type directly, ``Array`` is a static instance of 
    this type, so use the syntax as show in the examples below.

    If you don't specify this annotation, xlOil may still pass an array
    to your function if the user gives a range argument, e.g. A1:B2. In 
    this case you will get a 2-dim Array[object]. If you know the data 
    type you want, it is significantly more perfomant to specify it by
    annotation with this type.

    Examples
    --------
    The following shows the available options

        @xlo.func
        def array1(x: xlo.Array[int]):
            pass

        @xlo.func
        def array2(y: xlo.Array[float](dims=1)):
            pass

        @xlo.func
        def array3(z: xlo.Array[str](trim=False)):
            pass
    
    Methods
    -------

    **[]** : type    
        Types in square brackets are converted to numpy dtypes, so this
        means the only supported types are: int, float, bool, str, datetime, object.
        Numpy has a richer variety of dtypes than this but Excel does not. For the 
        float data type, xlOil will convert #N/As to numpy.nan but other values will 
        causes errors.

    **(dims=n)** : int    
        Arrays can be either 1 or 2 dimensional, 2 is the default.  Note the Excel has
        the following behaviour for writing arrays into an array formula range specified
        with Ctrl-Alt-Enter:
        "If you use a horizontal array for the second argument, it is duplicated down to
        fill the entire rectangle. If you use a vertical array, it is duplicated right to 
        fill the entire rectangle. If you use a rectangular array, and it is too small for
        the rectangular range you want to put it in, that range is padded with #N/As."

    **(trim=x)** : bool    
        By default xlOil trims arrays to the last row & column which contain a nonempty
        string or non-#N/A value. This is generally desirable, but can be disabled with 
        this paramter.

    """
    def __getitem__(self, elem_type=object):
        """
        Specifies a data type for the array. The syntax is::
            xlo.Array[float]
        """
        return _xloilArray(dtype=elem_type, shape=(1,1))


# Cheat to avoid needing Py 3.7+ for __class_getitem__
Array = ArrayType() 


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
        Parses a specified module looking for functions with with the xloil.func 
        decorator and register them. Does not search inside second level imports.

        The argument can be a module object, module name or path string. The module 
        is first imported if it has not already been loaded.
 
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

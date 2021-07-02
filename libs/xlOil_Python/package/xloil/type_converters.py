import inspect
import importlib.util
import typing
import numpy as np

from .shadow_core import Range, CellError, CannotConvert
from xloil import log

if importlib.util.find_spec("xloil_core") is not None:
    import xloil_core
    from xloil_core import (
        set_return_converter,
        CustomConverter as _CustomConverter,
        CustomReturn as _CustomReturn,
        Return_Cache as _Return_Cache,
        Return_SingleValue as _Return_SingleValue,
        Read_Array_object_2d as _Read_Array_object_2d,
    )
else:
    def _Return_Cache():
        pass
    def _Return_SingleValue():
        pass
    def _Read_Array_object_2d(trim):
        pass

"""
This annotation includes all the types which can be passed from xlOil to
a function. There is not need to specify it to xlOil, but it could give 
useful type-checking information to other software which reads annotation.
"""
ExcelValue = typing.Union[bool, int, str, float, np.ndarray, dict, list, CellError]

_READ_CONVERTER_PREFIX   = "Read_"
_RETURN_CONVERTER_PREFIX = "Return_"

def get_internal_converter(type_name, read_excel_value=True):
    """
    Attempt to find converter with standardised name like `Read_int`. Falls back
    to Read_cache if none found
    """
    to_from = _READ_CONVERTER_PREFIX if read_excel_value else _RETURN_CONVERTER_PREFIX
    name    = f"{to_from}{type_name}"
    found   = getattr(xloil_core, name, None)
    return None if found is None else found()

def _make_argconverter(base_type, impl, allow_range=False):
    # Only inheriting from one type is supported becasue inheriting from
    # certain type pairs e.g. (int, str) will give "instance lay-out conflict"
    # error and I don't know which pairs are impacted.  Afaik it's only
    # C-implemented types
    class ArgConverter(base_type):
        _xloil_converter = impl
        _xloil_allow_range = allow_range

        def __new__(cls, *args, **kwargs):
            """
            Allows return type converters to be "called" in the expected way.
            """
            return cls._xloil_converter.get_handler()(*args, **kwargs)

    return ArgConverter

def is_type_converter(obj):
    return hasattr(obj, "_xloil_converter")

def unpack_type_converter(obj):
    return obj._xloil_converter, obj._xloil_allow_range

def _make_metaconverter(base_type, converter_impl, create_wrapper, allow_range=False):

    if inspect.isclass(converter_impl):
        #
        # We use a metaclass approach to allow some linting with nice syntax: 
        # We want to write `def f(x: obj)` and `def g(x: obj(...))` which requires 
        # both of these expressions to evaluate to types identifiable to xloil as
        # a converter but also subclassing the conversion target to get type checking
        # and autocomplete. 
        #
        class MetaConverter(base_type):
                
                def __init__(self):
                    pass # Never called, but keeps linter quiet

                def __new__(cls, *args, **kwargs):
                    # Construct the inner class with the provided args
                    instance = converter_impl(*args, **kwargs)
                    # Embed it in a new Converter which inherits from target
                    return _make_argconverter(base_type, create_wrapper(instance), allow_range)

        # If the target obj has a no-arg constructor, we want to write:
        # `def fn(x: target_obj)`, so MetaConverter must be a valid 
        # converter itself. We try to insert the correct attributes
        try:
            MetaConverter._xloil_converter = create_wrapper(converter_impl())
            MetaConverter._xloil_allow_range = allow_range
        except TypeError:
            pass

        return MetaConverter
    else:
        return _make_argconverter(base_type, create_wrapper(converter_impl), allow_range)


def _make_tuple(obj):
    try:
        return tuple(obj)
    except TypeError:
        return obj, 


class _ArgConverters:
    
    _converters = dict()

    def add(self, converter, arg_type):
        """
        Registers a arg converter for a given type
        """
        internal = unpack_type_converter(converter)[0]
        log(f"Added arg converter for type {arg_type}", level='info')
        self._converters[arg_type] = internal
    
    def remove(self, arg_type):
        self._converters.remove(arg_type)

    def get_converter(self, arg_type):
        """
        Returns a _CustomConverter object which handles the given type or None
        """
        return self._converters.get(arg_type, None)

arg_converters = _ArgConverters()


def converter(to=typing.Callable, range=False, register=False):
    """
    Decorator which declares a function or a class to be a type converter.

    A type converter function is expected to take an argument of type:
    int, bool, float, str, ExcelArray, Range (optional)

    The type converter should return a python object, which could be an 
    ExcelArray or Range.

    A type converter class may take parameters into its constructor
    and hold state.  It should implement __call__ to handle type conversion.

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
    def decorate(impl):
        result = _make_metaconverter(to, impl, lambda x: _CustomConverter(x), range)

        if register:
            if to is not typing.Callable and is_type_converter(result):
                arg_converters.add(result, to)
            else:
                log(
                    f"Cannot register arg converter {impl.__name__}: requires a specifed 'to' type and a default-constructible instance",
                    level="warn")

        return result
    return decorate


# TODO: return converters should be able to register types
class _ReturnConverters:
    
    _converters = dict()
    _registered = False

    def add(self, converter, types):
        """
        Registers a return converter for a given single or iterable of types.
        """
        internal_converter = unpack_type_converter(converter)[0].get_handler()
        name = getattr(internal_converter, "__name__", type(internal_converter).__name__)
        log(f"Added return converter {name} for types {types}", level='info')
        try:
            for t in types:
                self._converters[t] = internal_converter # TODO: warn log on overwrite
        except TypeError:
             self._converters[types] = internal_converter
        
        # Register this object as the custom return converter tried by xlOil when 
        # a func does not specify its return type 
        if not self._registered:
            set_return_converter(_CustomReturn(self))
            self._registered = True

    
    def remove(self, return_type):
        self._converters.remove(return_type)
        if not any(self._converters):
            set_return_converter(None)
            self._registered = False

    def create_returner(self, return_type):

        """
        Creates a _CustomReturn object which handles the given type or returns None
        if no handlers can be found.  The _CustomReturn object is an internal xloil_core
        wrapper for a python-based return converter
        """

        found = self._converters.get(return_type, None)
        return _CustomReturn(found) if found is not None else None

    def __call__(self, obj):
        """
        Invoked by xlOil to try to convert the given object
        """
        for typ, converter in self._converters.items():
            try:
                if isinstance(obj, typ):
                    return converter(obj)
            except (CannotConvert):
                continue
        
        raise CannotConvert()


return_converters = _ReturnConverters()


def returner(types=None, register=False):

    """
    Decorator which declares a function or a class to be a return type converter.

    A return type converter should take a python object and return a type
    which xlOil knows how to return to Excel (which could be another type 
    converter).  It should raise ``CannotConvert`` if it cannot handle the
    given object.

    A return type converter class may take parameters into its constructor
    and hold state.  It should implement __call__ to handle type conversion

    Both functions and classes are turned into a class which inherit from 
    ``types`` (or Union[types] if more than one is provided).  This is to 
    support type hints only.

    If ``register`` is True, and one or more ``types`` specfied, the return
    converter is registered as a handler for those types, which means it 
    can be invoked when handling a ``func`` which does explicitly declare
    its return type.

    Examples
    --------
    
    ::

        @returner(MyType, register=True)
        def ReturnMyType(x):
            if isinstance(x, MyType):
                return x.__name__
            raise CannotConvert()

        @func
        def pyTest(x) -> MyType:
            return MyType()
            
    """

    def decorate(impl):

        # Multiple types can be specified, we just take the first one: even
        # this is unecessary, since xloil.funcs are never called from python
        # code (they can be, but it's unwise) their return type hints shouldn't 
        # be used by type checkers 
        try:
            from_type = next(iter(type))
        except TypeError:
            from_type = types

        result = _make_metaconverter(from_type, impl, lambda x: _CustomReturn(x))
        
        if register and types is not None:
            return_converters.add(result, types)

        return result

    return decorate

"""
Write `-> xloil.Cache` in a function declaration to force the output to be 
placed in the python object cache rather than attempting a conversion
"""
Cache = _make_argconverter(object, _Return_Cache())

"""
Use `-> xloil.SingleValue` in a function declaration to force the output to
be a single cell value. Uses the Excel object cache for returned arrays and 
the Python object cache for unconvertable objects
"""
SingleValue = _make_argconverter(object, _Return_SingleValue())

"""
The special AllowRange annotation allows functions to receive the argument
as an ExcelRange object if appropriate.  If a sheet reference (e.g. A1:B2) 
was not passed from Excel, xlOil converts as per ExcelValue.
"""
AllowRange = typing.Union[ExcelValue, Range]

class Array:
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

    **(dtype, dims, trim)** :    
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
 
    _xloil_converter = _Read_Array_object_2d(True)
    _xloil_allow_range = False

    def __new__(cls, dtype=object, dims=2, trim=True):
        name = f"{_READ_CONVERTER_PREFIX}Array_{dtype.__name__}_{dims or 2}d" 
        type_conv = getattr(xloil_core, name)(trim)
        return _make_argconverter(np.ndarray, type_conv, False)
     
import inspect
import importlib.util
import typing
import numpy as np
import functools

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
        Return_Array_object_2d as _Return_Array_object_2d,
    )
else:
    class _CustomConverter:
        def __init__(self, *arg, **kwargs):
            pass

    class _CustomReturn:
        def __init__(self, *arg, **kwargs):
            pass
        def get_handler(self):
            pass
    
    def set_return_converter(*arg, **kwargs):
        pass

    def _Return_Cache():
        pass
    def _Return_SingleValue():
        pass
    def _Read_Array_object_2d(trim):
        pass
    def _Return_Array_object_2d(trim):
        pass

"""
This annotation includes all the types which can be passed from xlOil to
a function. There is not need to specify it to xlOil, but it could give 
useful type-checking information to other software which reads annotations.
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

def _make_typeconverter(base_type, reader=None, writer=None, allow_range=False, source=None):
    # Only inheriting from one type is supported because inheriting from
    # certain type pairs e.g. (int, str) will give an "instance layout conflict"
    # error and I don't know which pairs are impacted.  Afaik it's only
    # C-implemented types
    class _TypeConverter(base_type):
        _xloil_arg_reader = (reader, allow_range) if reader is not None else None
        _xloil_return_writer = writer


        def __new__(cls, *args, **kwargs):
            """
            Allows return type converters to be "called" in the expected way.
            """
            return read(*args, **kwargs)

        @staticmethod
        def read(cls, value, *args, **kwargs):
            """
            Allows return type converters to be "called" in the expected way.
            """
            return cls._xloil_return_writer.get_handler()(*args, **kwargs)

    if source:
        functools.update_wrapper(_TypeConverter, source, updated=[])
        #_TypeConverter.__doc__      = source.__doc__
        #_TypeConverter.__name__     = source.__name__
        #_TypeConverter.__module__   = source.__module__
        #_TypeConverter.__qualname__ = source.__qualname__
    return _TypeConverter

def unpack_arg_converter(obj):
    return getattr(obj, "_xloil_arg_reader", None)

def unpack_return_converter(obj):
    return getattr(obj, "_xloil_return_writer", None)

def _make_metaconverter(base_type, impl, is_returner:bool, allow_range=False):

    if inspect.isclass(impl):
        #
        # We use a metaclass approach to allow some linting with nice syntax: 
        # We want to write `def f(x: obj)` and `def g(x: obj(...))` which requires 
        # both of these expressions to evaluate to types identifiable to xloil as
        # a converter but also subclassing the conversion target to get type checking
        # and autocomplete. 
        #
        class MetaConverter(base_type):
                
            #__doc__    = impl.__doc__
            #__name__   = impl.__name__
            #__qualname__ = impl.__qualname__
            #__module__ = impl.__module__

            def __init__(self):
                pass # Never called, but keeps linter quiet

            def __new__(cls, *args, **kwargs):
                # Construct the inner class with the provided args
                instance = impl(*args, **kwargs)
                # Embed it in a new Converter which inherits from target
                return _make_typeconverter(base_type, 
                                            _CustomConverter(instance.read) if hasattr(instance, "read") else None,
                                            _CustomReturn(instance.write) if hasattr(instance, "write") else None,
                                            allow_range)

        # If the target obj has a no-arg constructor, we want to write:
        # `def fn(x: target_obj)`, so MetaConverter must be a valid 
        # converter itself. We try to insert the correct attributes
        try:
            instance = impl()
            
            if hasattr(instance, "read"):
                MetaConverter._xloil_arg_reader = _CustomConverter(instance.read), allow_range
            if hasattr(instance, "write"):
                MetaConverter._xloil_return_writer = _CustomReturn(instance.write)

        except TypeError:
            pass

        functools.update_wrapper(MetaConverter, impl, updated=[])

        return MetaConverter
    elif is_returner:
        return _make_typeconverter(base_type, writer=_CustomReturn(impl), source=impl)
    else:
        return _make_typeconverter(base_type, _CustomConverter(impl), None, allow_range, source=impl)


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
        internal, _ = unpack_arg_converter(converter)
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


def converter(typeof=typing.Callable, range=False, register=False, direction="read"):
    """
    Decorator which declares a function or a class to be a type converter
    which serialises from/to a set of simple types understood by Excel and 
    general python types.

    A type converter class is expected to implement at least one of 
    ``read(self, val)`` and ``write(self, val)``. It may take parameters
    in its constructor and hold state. 

    A function is interpreted as a type reader or writer depending on the 
    *direction* parameter.

    **Readers**

    A *reader* converts function arguments to python types. It should receieve
    a value of: 

      *int*, *bool*, *float*, *str*, `xloil.ExcelArray`, `CellError`, `xloil.ExcelRange` (optional) 

    and return a python object or raise an exception (ideally `xloil.CannotConvert`).

    If ``range`` is True, xlOil may pass an *ExcelRange* or and *ExcelArray*
    object depending on how the function was invoked.  The converter should 
    handle both cases consistently.

    **Writers**

    A return type converter should take a python object and return a simple type
    which xlOil knows how to return to Excel. It should raise ``CannotConvert`` 
    if it cannot handle the given object.

    Parameters
    ----------

    typeof: 
        The type which the converter handles

    register: 
        If not True, registers the converter as a handler for ``typeof``, replacing
        any exsting handlers. For a reader, this means if ``typeof`` is used as an 
        argument annotation, this converter will be used.  For a writer, it enables 
        ``typeof`` as an return type annotation *and* it allows xlOil to try to call 
        this converter for Excel functions with no return annotation.
    
    range:
        For readers, setting this flag allows *xloil.ExcelRange* arguments to be passed

    direction:
        When decorating a function, the direction "read" or "write" determines the
        converter behaviour

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
        result = _make_metaconverter(typeof, impl, direction == "write", range)

        if bool(register) and typeof is not typing.Callable:

            global arg_converters, return_converters

            arg_converter = unpack_arg_converter(result)
            if arg_converter is not None:
                arg_converters.add(result, typeof)

            ret_converter = unpack_return_converter(result)
            if ret_converter is not None:
                try: 
                    # Is register an iterable of types?
                    return_converters.add(result, register)
                except TypeError:
                    pass
                return_converters.add(result, typeof)
            
        return result
    return decorate


class _ReturnConverters:
    
    _converters = dict()
    _registered = False

    def add(self, converter, types):
        """
        Registers a return converter for a given single or iterable of types.
        """
        # We unpack the python implementation of the converter so we can iterate 
        # through them trying to find one which doesn't raise CannotConvert
        converter_impl = unpack_return_converter(converter).get_handler()

        name = getattr(converter_impl, "__name__", type(converter_impl).__name__)
        log(f"Added return converter {name} for types {types}", level='info')

        try:
            for t in types:
                self._converters[t] = converter_impl # TODO: warn log on overwrite
        except TypeError:
            # Not iterable
            self._converters[types] = converter_impl
        
        # Register this singleton object as the custom return converter tried by xlOil 
        # when a func does not specify its return type 
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


def returner(typeof=None, register=False):
    """
    A proxy for converter(..., direction="write")
    """
    return converter(typeof, register=register, direction="write")


Cache = _make_typeconverter(object, writer=_Return_Cache())
"""
Using `-> xloil.Cache` in a function declaration to force the output to be 
placed in the python object cache rather than attempting a conversion
"""

SingleValue = _make_typeconverter(object, writer=_Return_SingleValue())
"""
Use `-> xloil.SingleValue` in a function declaration to force the output to
be a single cell value. Uses the Excel object cache for returned arrays and 
the Python object cache for unconvertable objects
"""

AllowRange = typing.Union[ExcelValue, Range]
"""
The special AllowRange annotation allows functions to receive the argument
as an ExcelRange object if appropriate.  If a sheet reference (e.g. A1:B2) 
was not passed from Excel, xlOil converts as per ExcelValue.
"""

class Array(np.ndarray):
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

    ::

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
        this paramter.  Has no effect when used for return values.

    """
 
    _xloil_arg_converter = _Read_Array_object_2d(True)
    _xloil_return_converter = _Return_Array_object_2d(False)
    _xloil_allow_range = False

    def __new__(cls, dtype=object, dims=2, trim=True, cache_return=False):
        name = f"{_READ_CONVERTER_PREFIX}Array_{dtype.__name__}_{dims or 2}d" 
        arg_conv = getattr(xloil_core, name)(trim)

        name = f"{_RETURN_CONVERTER_PREFIX}Array_{dtype.__name__}_{dims or 2}d" 
        return_conv = getattr(xloil_core, name)(cache_return)
        return _make_typeconverter(np.ndarray, arg_conv, return_conv, False)
     
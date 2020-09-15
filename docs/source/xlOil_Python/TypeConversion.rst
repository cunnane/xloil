============================
xlOil Python Type Conversion
============================

.. contents::
    :local:


Typing
------

xlOil reads ``typing`` annotations from decorated worksheet functions.  It supports
the following types:

   * bool
   * int
   * str
   * float
   * xloil.Array(dtype, dims, trim) (convert to a numpy array)
   * dict (expects a two-column array)
   * list (expects a 1-dim array)
   * xloil.PDFrame(...) (in xloil.pandas)

For example,

::

    @xlo.func
    def pySumNums(x: float, y: float, a: int = 2, b: int = 3) -> float:
        return x * a + y * b

Annotations which xlOil does not understand are ignored.


Cached Objects
--------------

If xlOil cannot convert a returned python object to Excel, it will place it in 
an object dictionary and return a reference string of the form
``<UniqueChar>[WorkbookName]SheetName!CellRef,#``

If a string of this kind if encountered when reading function arguments, xlOil 
automatically tries to fetch the corresponding python object. With this
mechanism you can pass python objects opaquely between functions by using
cache reference strings outputted from one function as arguments to another.

For example:

::

    @xlo.func
    def make_lambda(pow):
        return lambda x: x ** pow

    @xlo.func
    def apply_lambda(f, x):
        return f(x)

Since xlOil doesn't know how to convert a lambda function to an Excel object,
the first function will output a cache reference string.  That string will be 
automatically turned back into a lambda if passed as an argument to the second 
function.

The python cache is separate to the Core object cache accessed using `xloRef`
and `xloVal`.  The Core cache stores native Excel objects such as arrays.
When reading functions arguments xlOil tries to lookup strings in both of these
caches. 

The leading `UniqueChar` means xlOil can very quickly determine that a string
isn't a cache reference, so the overhead of checking if every string argument
is a cache object is very low in practice. 

Dates
-----

In the Python plugin, just applying the argument annotation `datetime` will request a date 
conversion. Dates returned from functions will be converted to Excel date numbers:

::

    from datetime import datetime, timedelta
    @func
    def AddDay(date: datetime):
        return date + timedelta(days = 1)


xlOil can interpret strings as dates. In the settings file, the key ``DateFormats`` 
specifies an array of date-formats to try when parsing strings. Naturally, adding more 
formats decreases performance.  The formats use the C++ ``std::get_time`` syntax,
see https://en.cppreference.com/w/cpp/io/manip/get_time.

Since ``std::get_time`` is **case-sensitive** on Windows, so is xlOil's date parsing
(this may be fixed in a future release as it is quite annoying for month names).

Excel has limited internal support for dates. There is no primitive date object in the  
XLL interface used by xlOil, but cells containing numbers can be formatted as dates.  
This means that  worksheet functions cannot tell whether values received are intended 
as dates - this applies to Excel built-in date functions as well: they will interpret 
any number as a date. (It is possible to check for date formatting via the COM interface 
but this would give behaviour inconsistent with the built-ins)

Excel does not understand timezones and neither does ``std::get_time``, so these
are currently unsupported.


Dicts and Keyword Arguments
---------------------------

If an xlOil decorated function has a ``**kwargs`` argument, it will expect the caller to
pass a two column array of key-value pairs, which it will convert to a dictionary.

It is also possible to return a ``dict`` from a function using the return type specifier.
Without specifying the return as dict, the default iterable converter would be used 
resulting in only the keys being output as an array.

The following example round-trips the provided keyword arguments:

::

    @xlo.func
    def pyTestKwargs(**kwargs) -> dict:
        return kwargs


Range Arguments
---------------

Range arguments allow a function to directly access a part of the worksheet. This 
allows macro functions to write to the worksheet or can be used for optimsation
if a function only requires a few values from a large input range.

A function can only receive range arguments if it is declared as *macro-type*. In 
addition, attempting to write to the worksheet during Excel's calculation cycle
will fail.

Annotating an argument with ``xlo.Range`` will tell xlOil to pass the function an
*ExcelRange* object, or fail if this is not possible.  An *ExcelRange* can only be
created when the input argument explicitly points to a part of the worksheet, not 
an array output from another function.

Annotating an argument with ``xlo.AllowRange`` will tell xlOil to pass an 
*ExcelRange* object if possible, otherwise fallback to the most appropriate basic
data type (int, str, array, etc.).


Custom Type Conversion
----------------------

A custom type converter is a callable which can create a python object
from a given *bool*, *float*, *int*, *str*, *ExcelArray* or *CellError*. 
Since each  argument to an Excel function can pass any of these types, the
converter must be able to handle all of them, or raise an exception. An 
*ExcelArray* represents an un-processed array argument - the type 
converter may opt to process only a part of this array for efficiency. 

The ``converter`` decorator tells xlOil that the following function or 
class is a type converter. Once decorated, the converter can be applied 
to an argument using the usual `typing` annotation syntax, or using the 
``args`` argument to ``@func()``.

By specifying ``xlo.converter(range=True)``, the type converter can opt to
receive *ExcelRange* arguments in addition to the other types.

::

    @xlo.converter()
    def arg_doubler(x):
        if isinstance(x, xlo.ExcelArray):
            x = x.to_numpy()
        return 2 * x

    @xlo.func
    def pyTestCustomConv(x: arg_doubler):
        return x


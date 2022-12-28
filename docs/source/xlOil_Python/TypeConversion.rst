============================
xlOil Python Type Conversion
============================

.. contents::
    :local:


Argument Types
--------------

xlOil function declarations in python look like:

::

    @xlo.func
    def DoSomething(x, y:float):
      return x

If no type is specified for an argument, xlOil will dynamically choose a type based
on the argument provied by Excel, this can be one of:

    * *bool*
    * *int*
    * *str*
    * *float*
    * *numpy.ndarray* (if an array or range is passed)
    * :py:class:`xloil.CellError`

Using ``typing`` annotations improves performance at the expense of static
typing.  Annonations also allow for user-defined conversion to any python type. 
xlOil has built-in support for the following annotations:

.. list-table:: Supported argument annotations
    :widths: 20 50
    :header-rows: 1

    * - Type
      - Comment
    * - *bool*
      - 
    * - *int*
      -
    * - *str*
      -
    * - *float*
      -
    * - *numpy.ndarray*
      - Use the :py:class:`xloil.Array` annotation rather than ndarray directly
    * - *dict*
      - Requires a 2-column input array. The first column is interpreted as keys
    * - *tuple*
      - Gives a tuple of tuple-of-tuples depending on number of input dimensions
    * - *datetime.date*
      - See :ref:`xlOil_Python/TypeConversion:Dates`
    * - *datetime.datetime*
      - See :ref:`xlOil_Python/TypeConversion:Dates`
    * - *pandas.DataFrame*
      - Can use the :py:class:`xloil.pandas.PDFrame` annotation for more conversion options. 
        Need to `import xloil.pandas` before use.
    * - *pandas.Timestamp*
      - Need to `import xloil.pandas` before use.
    * - :py:class:`xloil.Range`
      - See :ref:`xlOil_Python/TypeConversion:Range Arguments`
    * - :py:class:`xloil.AllowRange`
      - See :ref:`xlOil_Python/TypeConversion:Range Arguments` 
    * - <AnyType>
      - See :ref:`xlOil_Python/TypeConversion:Custom Type Conversion`

Annotations which xlOil does not understand are ignored.

Example:

::

    @xlo.func
    def pySumNums(x: float, y: float, a: int = 2, b: int = 3) -> float:
        return x * a + y * b


Return Types
------------

Like argument types, xlOil can read return type annotations. If no annotation
is specified xlOil tries the following conversions:
   
   * *None*
   * *int*
   * *float*
   * *numpy.ndarray*
   * *datetime*
   * :py:class:`xloil.CellError`
   * *str*
   * Registered custom return converters, see :ref:`xlOil_Python/TypeConversion:Custom Return Conversion`
   * iterable

If none of these succeeds, the object is placed in the cache, see :ref:`xlOil_Python/TypeConversion:Cached Objects`

.. list-table:: Supported return type annotations
    :widths: 20 50
    :header-rows: 1

    * - Type
      - Comment
    * - *bool*
      - 
    * - *int*
      -
    * - *str*
      -
    * - *float*
      -
    * - *numpy.ndarray*
      - Use the :py:class:`xloil.Array` annotation rather than ndarray directly
    * - *dict*
      - Outputs a 2-column array of key, value pairs
    * - *tuple*
      - A tuple of tuple-of-tuples produces a 1 or 2 dim array
    * - *datetime.date*
      - See :ref:`xlOil_Python/TypeConversion:Dates`
    * - *datetime.datetime*
      - See :ref:`xlOil_Python/TypeConversion:Dates`
    * - *pandas.DataFrame*
      - Can use the :py:class:`xloil.pandas.PDFrame` annotation for more conversion options. 
        Need to `import xloil.pandas` before use.
    * - *pandas.Timestamp*
      - Need to `import xloil.pandas` before use.
    * - *PIL.Image*
      - See :ref:`xlOil_Python/TypeConversion:Returning Images and Plots`
    * - *matplotlib.pyplot.Figure*
      - See :ref:`xlOil_Python/TypeConversion:Returning Images and Plots`
    * - :py:class:`xloil.Cache`
      - Placed the return value in the python object cache, see :ref:`xlOil_Python/TypeConversion:Cached Objects`.
    * - :py:class:`xloil.SingleValue`
      - Ensures the output will be a single cell value, not an array.
    * - <AnyType>
      - See :ref:`xlOil_Python/TypeConversion:Custom Return Conversion`

Cached Objects
--------------

If xlOil cannot convert a returned python object to Excel, it will place it in 
an object dictionary and return a reference string of the form

``<UniqueChar>SheetID!CellNumber,#``

xlOil automatically resolves cache string passed function arguments to their
objects.  With this mechanism you can pass python objects opaquely between 
functions.  You should not attempt to construct a cache string directly.

For example:

::

    @xlo.func
    def make_lambda(pow):
        return lambda x: x ** pow

    @xlo.func
    def apply_lambda(f, x):
        return f(x)

Since xlOil cannot convert a lambda function to an Excel object, it outputs a 
cache reference string.  That string is automatically turned back into a lambda 
if passed as an argument to the second function.

The python cache is separate to the Core object cache accessed using `xloRef`
and `xloVal`.  The Core cache stores native Excel objects such as arrays.
When reading functions arguments xlOil tries to lookup strings in both of these
caches. 

The leading *<UniqueChar>* means xlOil can very quickly determine that a string
isn't a cache reference, so the overhead of checking if every string argument
is a cache object is very low in practice. 

Dates
-----

Applying the argument annotation ``datetime.datetime`` requests a date conversion. Returning 
a ``datetime`` is allowed without a return annotation: the datetime will be converted to
an Excel date number:

::

    from datetime import datetime, timedelta
    @func
    def AddDay(date: datetime):
        return date + timedelta(days = 1)


xlOil can interpret strings as dates. In the settings file, the key ``DateFormats`` 
specifies an array of date formats to try when parsing strings. Naturally, adding more 
formats decreases performance. The formats use the C++ ``std::get_time`` syntax,
see https://en.cppreference.com/w/cpp/io/manip/get_time.

Since ``std::get_time`` is **case-sensitive** on Windows, so is xlOil's date parsing
(this may be fixed in a future release as it is quite annoying for month names).

Excel has limited internal support for dates. There is no primitive date object 
but cells containing numbers can be formatted as dates. This means that worksheet 
functions cannot tell whether numerical values are intended as dates - this applies
to Excel built-in date functions as well. (It is possible to check for date formatting
via the COM interface but this would give behaviour inconsistent with the built-ins)

Excel does not understand timezones and neither does ``std::get_time``, so these
are currently unsupported.


Dicts
-----

When the ``dict`` *argument type* annotation is specified, xlOil expects a two-column 
array of(*string*, *value*) to be passed.

Using a ``dict`` *return type* annotation allows a ``dict`` to be returned as as a 
two column array. Without the annotation, the default iterable converter would be invoked, 
resulting in only the keys being output.

Variable and Keyword Arguments
-------------------------------

If keyword args (`**kwargs`) are specified, xlOil expects a two-column array of 
(*string*, *value*) to be passed, the same as using a ``dict`` annotation. For variable
args (`*args`) xlOil adds a large number of trailing optional arguments. The variable
argument list is ended by the first missing argument.  If both *kwargs* and *args* are 
specified, their order is reversed in the Excel function declaration.

The following example shows dictionary and keyword aruments:

::

  @xlo.func
  def pyTestKwargs(lookup: dict, **kwargs) -> dict:
      lookup.update(kwargs)
      return lookup

The number of trailing optional arguments is limited by the maxiumum number of arguments 
allowed by Excel, which is 255 for a worksheet function and 60 for a local function.

Range Arguments
---------------

Range arguments allow a function to directly access a part of the worksheet. This 
allows macro functions to write to the worksheet or it can be used for optimisation
if a function only requires a few values from a large input range.

A function can only receive range arguments if it is declared as *macro-type*. In 
addition, attempting to write to a Range during Excel's calculation cycle will fail.

Annotating an argument with :py:class:`xlo.Range` will tell xlOil to pass the function an
:py:class:`Range` object, or fail if this is not possible.  An :py:class:`Range` 
can only be created when the input argument explicitly points to a part of the worksheet, not 
an array output from another function.

Annotating an argument with :py:class:`xlo.AllowRange` will tell xlOil to pass an 
:py:class:`Range` object if possible, otherwise one of the other basic data types
(int, str, array, etc.).


Custom Type Conversion
----------------------

A custom type converter is a function or a class which serialises between a set 
of simple types understood by Excel and general python types.

A type converter class is expected to implement at least one of ``read(self, val)`` 
and ``write(self, val)`` and be decorated with :py:func:`xloil.converter`.
It may take parameters in its constructor and hold state. 

A function can be interpreted as a type reader or writer depending on the parameters
passed to the :py:func:`xloil.converter` decorator.

The ``read(self, val)`` method or a function decorated as a reader or argument converter 
should be able to accept a value of: 

    *int*, *bool*, *float*, *str*, :py:class:`xloil.ExcelArray`, :py:class:`CellError`, 
    :py:class:`xloil.Range` (optional) 

and return a python object or raise an exception (ideally :py:class:`xloil.CannotConvert`).

An :py:class:`xloil.ExcelArray` represents an un-processed array argument, a
handle to the raw Excel object not yet converted to a *numpy* array.  The converter
may opt to process only a part of this array for efficiency. 

A converter may be used by name in *typing* annotations for :py:func:`xloil.func` 
functions.  In addition, the converter can register as the handler for a specific type 
which enables that type to be used in annotations.  For registration, the converter must
be default-constructible (or be a function).

By decorating with ``@xloil.converter(range=True)``, the type converter can opt to
receive :py:class:`Range` arguments in addition to the other types.


::

    @xlo.converter()
    def arg_doubler(x):
      if isinstance(x, xlo.ExcelArray):
        x = x.to_numpy()
      return 2 * x

    @xlo.func
    def pyTestCustomConv(x: arg_doubler):
      return x

    @xlo.converter(typeof=bytes, register=True)
    class StrToBytes:
      def __init__(self, encoding='utf-8'):
        self._encoding = encoding
      def read(self, val):
        return val.encode(self._encoding)
      def write(self, val):
        return val.decode(self._encoding)
      
    @xlo.func
    def Pad(text: bytes, size: int) -> StrToBytes('utf-8'):
      return text.center(size) 

Custom Return Conversion
------------------------

A return type converter should take a python object and return a simple type
which xlOil knows how to return to Excel. It should raise :py:class:`xloil.CannotConvert` 
if it cannot handle the given object.

It can be a class implementing ``write(self, val)`` and decorated with 
:py:class:`xloil.converter` or a function decorated with :py:class:`xloil.returner`
or :py:class:`xloil.converter`.

A return converter can register as the handler for a specific type which enables that 
type to be used in return annotations *and* allows xlOil to try to call 
the converter for Excel functions with no return annotation, see :ref:`xlOil_Python/TypeConversion:Return Types`.
        

::

    @xlo.returner(typeof=MyType, register=True)
    def mytypename(val):
      return val.__name__
    
    @xlo.func
    def MakeMyType():
      return MyType()
  

Returning Images and Plots
--------------------------

By using custom return converters you can return `PIL` or `pillow` image 
objects from worksheet functions. The returned image can be automatically 
sized to the calling range, or any offset from it, but it floats like a 
normal picture in Excel.  Calling the worksheet function again removes
the previous image and replaces it with a new one.

::

    import xloil.pillow
    from PIL import Image
    
    @xlo.func(macro=True) # macro permissions required
    def ShowPic(filename):
        return Image.open(filename)


Importing ``xloil.pillow`` registers a custom return converter for ``PIL.Image``.
To gain control over the image size and position, use the :py:class:`xloil.pillow.ReturnImage`
return annotation.

Similarly a matplotlib figure can be returned directly

::

    import xloil.matplotlib

    @func(macro=True)
    def Plot(x, y):
        fig = pyplot.figure(figsize=(5,5))
        fig.add_subplot(111).plot(x, y)
        return fig

Importing ``xloil.matplotlib`` registers a custom return converter for 
``matplotlib.pyplot.Figure``. To gain control over the plot size and position, 
use the :py:class:`xloil.matplotlib.ReturnFigure` return annotation.

Both of these converters use :py:class:`xloil.insert_cell_image`.

=============================
xlOil Python Package Support
=============================

xlOil has built-in converters for several external libraries

.. contents::
    :local:


Pandas
------

Pandas support can be enabled in *xlOil* with:

::

    import xloil.pandas

This registers a return converter which allows pandas dataframes to be returned
from Excel functions in a natural table format.  Not if a function always returns 
a dataframe it is more performant to specify  ``pandas.DataFrame`` as the return type
annotation.

The annotation ``pandas.DataFrame`` can also be used for arguments. The converter by 
default expects a two dimensional array with headings in the first row.

To gain more control over the *DataFrame* conversion, use the annotion 
`xloil.pandas.PDFrame`.  Examples:

::

    @xlo.func
    def getDataField(data: PDFrame(headings=True), columns) -> PDFrame(headings=False):
        return data[columnName]

    @xlo.func
    def readDatedData(df: PDFrame(headings=True, index=['Date', 'Type'], dates=['Date'])):
        # df will have a pandas.MultiIndex index
        return df


The parameters have slightly different interpretations when reading or outputting arguments.

The `headings`` parameter when reading, if True, interprets the first row as column heading or 
if it is an int inperets the first *N* rows as a *pandas.MultiIndex* heading.  When outputting, 
if it is True/False outputs column headings or not.

The `index`` parameter when reading specifies column(s) which should be set as the index: xloil 
calls `DataFrame.set_index(<index>)`, so a column name or list-like of columns names can be 
given.  When writing: if *index* is set to False, the index is not output, otherwise it is.

The `dates` parameters specifies which columns to convert from from Excel serial date numbers 
to *numpy.datetime64*. Since Excel stores dates as floating points it is not possible for
xlOil to identify date columns automatically.  This parameter has no effect when outputting as 
*datetime* arrays are converted automatically.

We can call the *PDFrame* converter in the function itself if we want to control the arguments
passed to it based on the function's inputs (this isn't possible if it is used as a decorator).

::

    @xlo.func
    def dFrameMultiHeadings(
            df: PDFrame(headings=2),
            outputHeadings=False):
        return PDFrame(headings=outputHeadings)(df)


See :doc:`Example` for more examples.

Matplotlib
----------

Importing ``xloil.matplotlib`` defines a return converter so matplotlib figures
can be returned from worksheet functions.  By default they are resized to the cell.
Subseqent figures returned from the same cell overwrite previous ones.
Returning a figure requires setting ``macro=True`` in the :obj:`xloil.func` declaration.

::

    import xloil.matplotlib
    from matplotlib import pyplot
    
    @xlo.func(macro=True)
    def pyTestPlot(x, y, **kwargs):
        fig = pyplot.figure()
        ax.plot(x, y, **kwargs)
        return fig

    @xlo.func(macro=True)
    def pyTestPlot(x, y, width, height, **kwargs):
        fig = pyplot.figure()
        ax.plot(x, y, **kwargs)
        return xloil.matplotlib.ReturnFigure(size=(width, height), pos='top')(fig)

PIL / pillow
------------

Importing ``xloil.pillow`` defines a return converter so PIL images
can be returned from worksheet functions.  By default they are resized to the cell.
Subseqent images returned from the same cell overwrite previous ones.
Returning an image requires setting ``macro=True`` in the :obj:`xloil.func` declaration.

::

    import xloil.pillow
    from PIL import Image

    @xlo.func(macro=True)
    def pyTestPic():
        im = Image.open("MyPic.jpg")
        return im

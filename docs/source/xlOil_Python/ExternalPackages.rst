=================================
xlOil External Package Support
=================================

xlOil has built-in converters for several external libraries

.. contents::
    :local:


Pandas
------

Pandas support can be enabled in *xlOil* with:

::

    import xloil.pandas

This registers a return converter which allows pandas dataframes to be returned
from Excel functions in a table format.  If the function always returns a
dataframe it is more performant to specify  ``pandas.DataFrame`` as the return type
annotation.

It also defines an argument converter for the annotation ``pandas.DataFrame``. This
expects a two dimensional array with headings in the first row.

To gain more control over the DataFrame conversion, use the annotion 
`xloil.pandas.PDFrame`.

For example,

::

    @xlo.func
    def getDataField(data: PDFrame(headings=True), columnName):
        return data[columnName]

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

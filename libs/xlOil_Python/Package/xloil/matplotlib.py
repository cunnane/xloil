try:
    import matplotlib
    matplotlib.use('Agg')
    from matplotlib import pyplot
except ImportError:
    from ._core import XLOIL_READTHEDOCS
    if XLOIL_READTHEDOCS:
        class pyplot:
            class Figure:
                # Placeholder for matplotlib.pyplot.Figure
                ...

from xloil import *

@func(macro=True)
def xloPyPlot(x, y, width:float=None, height:float=None, **kwargs):
    fig = pyplot.figure()
    if width is not None:
        fig.set_size_inches(width, height)
    ax = fig.add_subplot(111)
    ax.plot(x, y, **kwargs)
    return fig


@returner(target=pyplot.Figure, register=True)
class ReturnFigure:
    """
        Inserts a plot as an image associated with the calling cell. A second call
        removes any image previously inserted by the same calling cell.

        Parameters
        ----------

        size:  
            * A tuple (width, height) in points. 
            * "cell" to fit to the caller size
            * "img" or None to keep the original image size
        pos:
            A tuple (X, Y) in points. The origin is determined by the `origin` argument
        origin:
            * "top" or None: the top left of the calling range
            * "sheet": the top left of the sheet
            * "bottom": the bottom right of the calling range
    """

    def __init__(self, size=None,  pos=(0, 0), origin:str=None):
        self._shape = (size, pos, origin)

    def write(self, fig:pyplot.Figure):
        return insert_cell_image(lambda filename: fig.savefig(filename, format="png"), *self._shape)

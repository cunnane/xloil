import matplotlib
matplotlib.use('Agg')
from matplotlib import pyplot
from .xloil import *


@func(macro=True)
def xloPyPlot(x, y, width:float=None, height:float=None, **kwargs):
    fig = pyplot.figure()
    if width is not None:
        fig.set_size_inches(width, height)
    ax = fig.add_subplot(111)
    ax.plot(x, y, **kwargs)
    return fig


@returner(types=pyplot.Figure, register=True)
class ReturnFigure:
    def __init__(self, size=None,  pos=(0, 0), origin:str=None):
        self._shape = (size, pos, origin)

    def __call__(self, fig:pyplot.Figure):
        return insert_cell_image(lambda filename: fig.savefig(filename, format="png"), *self._shape)

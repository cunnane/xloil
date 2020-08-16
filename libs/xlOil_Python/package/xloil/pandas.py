import pandas as pd
import numpy as np
from .xloil import *
import typing

@converter(pd.DataFrame)
class PDFrame:
    """
    Converter which takes tables with horizontal records to pandas dataframes.

    Examples
    --------

    ::

        @xlo.func
        def array1(x: xlo.PDFrame(int)):
            pass

        @xlo.func
        def array2(y: xlo.PDFrame(float, headings=True)):
            pass

        @xlo.func
        def array3(z: xlo.PDFrame(str, index='Index')):
            pass
    
    Methods
    -------

    **PDFrame(element, headings, index)** : 
            
        element : type
            Pandas performance can be improved by explicitly specifying  
            a type. In particular, creation of a homogenously typed
            Dataframe does not require copying the data.

        headings : bool
            Specifies that the first row should be interpreted as column
            headings

        index : various
            Is used in a call to pandas.DataFrame.set_index()

    """
    def __init__(self, element=None, headings=True, index=None):
        # TODO: use element_type!
        self._element_type = element
        self._headings = headings
        self._index = index

    def __call__(self, x):
        if isinstance(x, ExcelArray):
            df = None
            idx = self._index
            if self._headings:
                if x.nrows < 2:
                    raise Exception("Expected at least 2 rows")
                headings = x[0,:].to_numpy(dims=1)
                data = {headings[i]: x[1:, i].to_numpy(dims=1) for i in range(x.ncols)}
                if idx is not None and idx in data:
                    index = data.pop(idx)
                    df = pd.DataFrame(data, index=index).rename_axis(idx)
                    idx = None
                else:
                    # This will do a copy.  The copy can be avoided by monkey
                    # patching pandas - see stackoverflow
                    df = pd.DataFrame(data)
            else:
                df = pd.DataFrame(x.to_numpy())
            if idx is not None:
                df.set_index(idx, inplace=True)
            return df
        
        raise Exception(f"Unsupported type: {type(x)!r}")


def PandasReturn(val):
    if type(val) is pd.DataFrame:
        # TODO: not exactly performant!
        header = val.columns.values
        index = val.index.values[:,np.newaxis]
        pad = np.full((np.atleast_2d(header).shape[0], index.shape[1]), ' ', dtype=object)
        pad[0, 0] = val.index.name
        return np.block([[pad, header], [index, val.values]])
    elif type(val) is pd.Timestamp:
        return val.to_pydatetime()
    else:
        raise CannotConvert()


return_converters.add(PandasReturn)
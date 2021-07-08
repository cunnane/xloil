import pandas as pd
import numpy as np
from xloil import *
import typing

@converter(pd.DataFrame, register=True)
class PDFrame:
    """
    Converter which takes tables with horizontal records to pandas dataframes.

    **PDFrame(element, headings, index)**

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
    
    Parameters
    ----------
        
    element : type
        Pandas performance can be improved by explicitly specifying  
        a type. In particular, creation of a homogenously typed
        Dataframe does not require copying the data. Not currently
        implemented!

    headings : bool
        Specifies that the first row should be interpreted as column
        headings

    index : various
        Is used in a call to pandas.DataFrame.set_index()

    """
    def __init__(self, element=None, headings=True, index=None):
        # TODO: use element_type in the dataframe construction
        self._element_type = element
        self._headings = headings
        self._index = index

    def read(self, x):
        # A converter should check if provided value is already of the correct type.
        # This can happen as xlOil expands cache strings before calling user converters
        if isinstance(x, pd.DataFrame):
            return x

        elif isinstance(x, ExcelArray):
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
        
        raise CannotConvert(f"Unsupported type: {type(x)!r}")

    def write(self, val):
        # Construct this array
        #   [filler]      [col_labels]
        #   [row_labels]  [values]

        row_labels = val.index.values[:, np.newaxis]

        if self._headings:
            col_labels = val.columns.values
            filler_size = (np.atleast_2d(col_labels).shape[0], row_labels.shape[1])
            filler = np.full(filler_size, ' ', dtype=object)
        
            # Write the name of the index in the top left
            filler[0, 0] = val.index.name
       
            return np.block([[filler, col_labels], [row_labels, val.values]])
        else:
            return np.block([[row_labels, val.values]])

@converter(typeof=pd.Timestamp, register=True)
class PandasTimestamp:
    """
        There is not need to use this class directly in annotations, rather 
        use ``pandas.Timestamp``
    """

    def read(self, val):
        return pd.Timestamp(val)

    def write(self, val):
        return val.to_pydatetime()

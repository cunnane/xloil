try:
    import pandas as pd
    from pandas.api.types import is_datetime64_any_dtype, is_numeric_dtype
except ImportError:
    from ._core import XLOIL_READTHEDOCS
    if XLOIL_READTHEDOCS:
        class pd:
            class DataFrame:
                # Placeholder for pandas.DataFrame
                ...

            class Timestamp:
                # Placeholder for pandas.Timestamp
                ...

from dataclasses import dataclass
from re import I
import numpy as np
from xloil import *
import typing

@converter(pd.DataFrame, register=True)
class PDFrame:
    """
    Converts beteeen data tables with horizontal records and *pandas* DataFrames.

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
        

    headings: bool / int
        When reading: if True, interprets the first row as column headings, if
        an int inprets the first *N* rows as a *MultiIndex* heading,
        When writing: if True, outputs column headings

    index: bool / index-spec
        When reading: specifies column(s) which should be treated as the index: 
        xloil calls `DataFrame.set_index(<index>)`, so a column name or list-like
        of columns names can be given
        When writing: if explicitly set to False, the index is not output. Any
        other value causes the index to be output

    dates: list
        When reading: attempt to convert the named columns from Excel serial 
        date numbers to numpy datetime.

    allow_object: bool (default False)
        When writing, if False, any non-numeric objects are converted to string.
        This prevents a large number of object refs being created.

    dtype: type
        Not currently implemented!

    """
    def __init__(self, headings=True, index=None, allow_object=False, dates=None, dtype=None):
        # TODO: use element_type in the dataframe construction
        self._element_type = dtype
        self._headings = headings
        self._index = index
        self._allow_object = allow_object
        self._parse_dates = dates

    def read(self, x):
        # A converter should check if provided value is already of the correct type.
        # This can happen as xlOil expands cache strings before calling user converters
        if isinstance(x, pd.DataFrame):
            return x

        elif isinstance(x, ExcelArray):

            data = {i: x[1:, i].to_numpy(dims=1) for i in range(x.ncols)}
            # This will do a copy.  The copy can be avoided by monkey
            # patching pandas - see stackoverflow
            df = pd.DataFrame(data, copy=False)

            if self._headings is True or self._headings == 1:
                if x.nrows < 2:
                    raise Exception("Expected at least 2 rows")
                headings = x[0,:].to_numpy(dims=1)
                df.set_axis(headings, axis=1, inplace=True)

            elif self._headings > 1:
                if x.nrows < 1 + self._headings:
                    raise Exception(f"Expected at least {1 + self._headings} rows")
                headings = x[:self._headings,:].to_numpy(dims=2)
                df.set_axis(pd.MultiIndex.from_arrays(headings), axis=1, inplace=True)

            if self._parse_dates is not None:
                for col in self._parse_dates:
                    if col in df:
                        df[col] = to_datetime(df[col].values)

            if self._index is not None:
                df.set_index(self._index, inplace=True)

            return df
        
        raise CannotConvert(f"Unsupported type: {type(x)!r}")

    def _to_array(data):
        return data.values if (
            self._allow_object 
            or is_datetime64_any_dtype(data) 
            or is_datetime64_any_dtype(data)
        ) else data.astype(str).values

    def write(self, frame: pd.DataFrame):

        import xloil_core

        columns = [self._to_array(frame[col]) for col in frame]

        if self._index is not False:
            index = [
                self._to_array(frame.index.get_level_values(i)) 
                for i in range(frame.index.nlevels)
                ]
        else:
            index = None

        if self._headings:
            headings = [
                self._to_array(frame.columns.get_level_values(i))
                for i in range(frame.columns.nlevels)
                ]
        else:
            headings = None

        return xloil_core._table_converter(
            frame.shape[1], 
            frame.shape[0],
            columns=columns,
            index=index,
            index_name=frame.index.names,
            headings=headings)


@converter(target=pd.Timestamp, register=True)
class PandasTimestamp:
    """
        There is not need to use this class directly in annotations, rather 
        use ``pandas.Timestamp``
    """

    def read(self, val):
        return pd.Timestamp(val)

    def write(self, val):
        return val.to_pydatetime()

try:
    import pandas as pd

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


from xloil import converter, to_datetime, ExcelArray, CannotConvert
import numpy as np
from collections.abc import Iterable

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

    cache_objects: bool (default False)
        When writing, if False, any objects which cannot be convertered by a
        known converter are converted to string via `str(x)`. This prevents a 
        large number of (possibly unhelpful) object refs being created.

    dtype: type
        Not currently implemented!

    """
    def __init__(self, headings=True, index=None, cache_objects=False, dates=None, dtype=None):
        # TODO: use element_type in the dataframe construction
        self._element_type = dtype
        self._headings = headings
        self._index = index
        self._cache_objects = cache_objects
        self._parse_dates = dates

    def read(self, x):
        # A converter should check if provided value is already of the correct type.
        # This can happen as xlOil expands cache strings before calling user converters
        if isinstance(x, pd.DataFrame):
            return x

        elif isinstance(x, ExcelArray):

            n_headings = int(self._headings)
            if x.nrows < n_headings:
                raise ArgumentError(f"Expected at least {n_headings} rows")

            data = {i: x[n_headings:, i].to_numpy(dims=1) for i in range(x.ncols)}
            # This will do a copy.  The copy can be avoided by monkey
            # patching pandas - see stackoverflow
            df = pd.DataFrame(data, copy=False)

            if n_headings == 1:
                headings = x[0,:].to_numpy(dims=1)
                df.set_axis(headings, axis=1, inplace=True)

            elif n_headings > 1:
                headings = x[:n_headings,:].to_numpy(dims=2)
                df.set_axis(pd.MultiIndex.from_arrays(headings), axis=1, inplace=True)

            if self._parse_dates is not None:
                for col in self._parse_dates:
                    if col in df:
                        df[col] = to_datetime(df[col].values.ravel())

            if self._index is not None:
                df.set_index(self._index, inplace=True)

            return df
        
        raise CannotConvert(f"Unsupported type: {type(x)!r}")

    def write(self, frame: pd.DataFrame):

        import xloil_core

        columns = [col.values for _, col in frame.items()]

        # If outputting the index, we prepare an array for each index level
        if self._index is not False:
            index = [
                frame.index.get_level_values(i).values
                for i in range(frame.index.nlevels)
                ]
        else:
            index = None

        # If outputting the columns, we prepare an array for each column level
        if self._headings:
            headings = [
                frame.columns.get_level_values(i).values
                for i in range(frame.columns.nlevels)
                ]
        else:
            headings = None

        # The index names may be a list of tuple (if index was created from a multi column 
        # index dataframe) or strings or None if no name was given.  We form this into an 
        # array of size column_levels x index_levels
        index_names = np.empty((frame.columns.nlevels, frame.index.nlevels), dtype=object)
        for j, name in enumerate(frame.index.names):
            if isinstance(name, Iterable) and len(name) <= index_names.shape[0]:
                for i, x in enumerate(name):
                    index_names[i, j] = x
            elif name is not None:
                index_names[0, j] = name
            
        return xloil_core._table_converter(
            frame.shape[1], 
            frame.shape[0],
            columns=columns,
            index=index,
            index_name=index_names.ravel(),
            headings=headings,
            cache_objects=self._cache_objects)


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

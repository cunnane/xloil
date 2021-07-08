=================================
xlOil Python Pandas Support
=================================

Pandas support can be enabled in *xlOil* with:

::

    import xloil.pandas

This registers a return converter which allows pandas dataframes to be returned
from Excel functions in a table format.  If the function always returns a
dataframe it is more performant to specify  ``pd.DataFrame`` as the return type
annotation.

It also defines an argument converter for the annotation ``pd.DataFrame``. This
expects a two dimensional array with headings in the first row.

To gain more control over the DataFrame conversion, use the annotion 
`xloil.pandas.PDFrame`.

For example,

::

    @xlo.func
    def getDataField(data: PDFrame(headings=True), columnName):
        return data[columnName]

See :doc:`Example` for more examples.

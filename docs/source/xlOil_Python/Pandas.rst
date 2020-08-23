=================================
xlOil Python Pandas Support
=================================

Pandas support can be enabled in *xlOil* with:

::

    import xloil.pandas

This will allow pandas dataframes to be returned from Excel functions in a table format. 

To automatically convert a function argument to a dataframe mark it as type
`xloil.pandas.PDFrame`.  For example,

::

    @xlo.func
    def getDataField(data: PDFrame(headings=True), columnName):
        return data[columnName]

See :doc:`Example` for more examples.

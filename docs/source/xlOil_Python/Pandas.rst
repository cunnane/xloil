=================================
xlOil Python Pandas Support
=================================

Pandas support can be enabled in *xlOil* with:

::

    import xloil.pandas

This will allow pandas dataframes to be returned from Excel functions in a table format. 
To automatically convert a function argument to a dataframe mark it as type
`xloil.pandas.PDFrame`. See :doc:`xlOil_Python_Example` for an example of this in action.
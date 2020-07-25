===========
xlOil Utils
===========

The Utils plugin contains general purpose tools to manipulate data,
particularly arrays.

.. contents:: Contents
	:local:
    
xloBlock: creates an array from blocks
----------------------------------------

.. function:: xloBlock(layout, arg1, arg2, arg3,...)

    Creates a block matrix given a layout specification.

    Layout has the form `1, 2, 3; 4, 5, 6` where numbers refer to
    argument numbers provided (note indexing is 1-based). Commas
    divide argument numbers, semi-colons indicate a new row. xlOil 
    expands the blocks row by row.

    Omiting the argument number, i.e. two consecutive commas gives an
    auto-sized padded space. You can only have one per row.

    Whitespace is ignored in the layout specification.

    Arguments may be an array or a single value, which is interpreted as
    a 1x1 array.

    Any holes in the result are filled with #N/A - this is preferrable 
    over Excel's 'empty' value which is transformed to a zero when 
    written to the sheet.

    **Examples**

    ::

        =xloBlock("1, 2; 3, 4", A A, B B, C C, D D)
                                A A  B B

        --> A A B B
            A A B B
            C C D D


        =xloBlock("1, 2; 3,, 4", A A, B B, C, D)
                                 A A  B B

        --> A A B B
            A A B B
            C - - D

xloFill: gives an array filled with a single value
--------------------------------------------------

.. function:: xloFill(value, numRows, numColumns)

    Creates an array of size numRows by numColumns filled with value.

    Throws if value is not a single value.

xloFillNA: replaces Empy or #N/A with a specified value
----------------------------------------------------------

.. function:: xloFillNA(value, array, [trim])

    Replaces any #N/A or Empty with a specified value. #N/As in Excel 
    result from holes, e.g. from xloBlock or array which does not fill
    the space allocated in an array formula.  Empty values are particularly
    problematic: they correpond to Excel reading an empty cell in an array
    but are transformed to a zero when written to the sheet.

    Throws if value is not a single value.

    The optional trim parameter can be used to disable xlOil's default
    array-trimming behaviour of resizing to the last non-empty, non-NA 
    row and column.


xloSort: sorts an array by multiple columns
----------------------------------------------------------

.. function:: xloSort(Array, Order, [colOrHeading1], [colOrHeading2], ...)

    Sorts data by one or more column keys. The function behaves similarly
    to Excel's sort command, but works as a sheet function.

    The `Order` parameter should be a string with a descriptor character
    for each column to used as a sort key.  The character described how 
    the column data will be compared. Allowed characters and their meaning 
    are:

        *a*: ascending

        *A*: ascending case-sensitive,

        *d*: descending

        *D*: descending case-sensitive

        *whitespace*: ignored

    Each subsequent argument should be the (1-based) number of a column
    or string.  If any strings are specified, the first row of `Array`
    is interpreted as column headers and the strings are matched against
    these headers.

    The order of the column specifiers indicates the prescendence in the
    sort order.

    xloSort works in-place for speed but uses Excel's variant types. This 
    means it will not be 

    **Examples**

    ::

        =xloSort( { Baz    4 } , "a d", 2, 1)
                  { Bar    4 }              
                  { Boz    2 }             

        -> 2    Boz
           4    Baz
           4    Bar    



xloPad: pads an array to have at least two rols and columns
-----------------------------------------------------------

.. function:: xloPad(array)

    xloPad works around the Excel 'feature' that duplicates arrays
    with only one row or column when they are used in an array formula with 
    a larger target size.  This can be frustrating for display of variable 
    sized array results.  xloPad simply ensures the array has at least two
    rows and columsn, adding #N/A where required to fill the space.


xloConcat: concatenates strings
--------------------------------

.. function:: xloConcat([separator], valOrArray1, valOrArray2, ...)

    Concatenates strings or other values with an optional separator.
    Concatenation is in argument order. The separator may be blank or any
    value which can be converted to a string.  Non string arguments are 
    converted to string before concatenation.  Arrays are concatenated
    along rows using the sparator if specified.


xloSplit: splits strings at a separator
---------------------------------------

.. function:: xloSplit(stringOrArray, separators, [consecutiveAsOne])

    Splits a string at one or more separator characters, returning an array.
    A separators can only be a single character, but multiple separators
    can be specified. If `consecutiveAsOne` is omitted or TRUE, consecutive
    separators are treated as one, otherwise they generate empty cells.

    If a array of strings is passed, splitting will occur on each array
    element and the array orientation wil be preserved. The array must be 
    1-dimensional

    Any non string values are ignored - no coercision is performed.

    The `separators` input can be a string containing any number of characters;
    each will be treated as a distinct separator - multi-character separators
    are not supported.

    **Examples**

    ::

        =xloSplit("Foo:Bar,,Baz", ":,")

        -> Foo      Bar     Baz


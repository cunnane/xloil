===========
xlOil Utils
===========

The utils plugin contains general utils to manipulate arrays in Excel

.. contents::
	:local:
    
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


.. function:: xloFill(value, numRows, numColumns)

    Creates an array of size numRows by numColumns filled with value.

    Throws if value is not a single value.

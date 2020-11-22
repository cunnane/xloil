======================
XLL Special Arguments
======================

User-defined Excel functions in XLLs usually take a number of `const ExcelObj&` arguments
and return a `ExcelObj*`.  xlOil supports a number of other possibilities, the most useful
of which are `RangeArg` and `AsyncHandle`.

RangeArg
--------

Declaring a class as `const RangeArg&` instead of an `ExcelObj` in registered function tells 
xlOil to allow range references to be passed.  These are sheet addresses which avoid copying
the underlying data.  Normally, references to areas on the sheet are converted to arrays
before being passed to the function.  A `RangeArg` only allows the possibility that a range
reference be specified, `RangeArg` inherits from `ExcelObj` and so may actually be any Excel 
type.

AsyncHandle
-----------

Declaring an argument as `const AsyncHandle&` causes the function to be declared as native
async type to Excel. Such a function does not return a value directly but invokes the 
`returnValue` method on its handle to pass a value to Excel (which should only be done once).

ExcelObj pointer
----------------

You may opt to receive a `const ExcelObj*` instead of a reference. This is purely preference
and has no other impact.

FPArray
-------

An `FPArray` argument is a two-dimensional array of double. This is naturally very fast to 
access, particulary when combined with inplace returns but has a significant drawback: if
if any value in the array passed to the function is not a number, Excel will return *#VALUE!*
without actually invoking the function.

Inplace returns
----------------

Excel supports returning values by modifying arguments in place.  This is most useful for the 
FPArray type.  To declare inplace return, simply make the argument a non-const ref or pointer.

It is also possible to returning an `ExcelObj` in-place but this is disabled 
by default. In the words of the XLL SDK:


    "Excel permits the registration of functions that return an XLOPER by modifying 
    an argument in place. However, if an XLOPER argument points to memory, and the 
    pointer is then overwritten by the return value of the DLL function, Excel can 
    leak memory. If the DLL allocated memory for the return value, Excel might try 
    to free that memory, which could cause an immediate crash.  Therefore, you should 
    not modify XLOPER/XLOPER12 arguments in place."


In practice, it can be safe to modify an ExcelObj in place, for instance *xloSort*
modifies its input by changing the row order in the array, but without changing memory 
allocation.  However it does not use inplace return becasue the general difficulty with 
this technique is that you cannot return a type different to the one passed, in particular
you cannot return an error message or error type. 

To enable inplace `ExcelObj` returns, define the macro 
`XLOIL_UNSAFE_INPLACE_RETURN` before including any xlOil headers.
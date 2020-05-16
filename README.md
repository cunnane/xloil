
xlOil
=====

xlOil is a framework for building Excel language bindings. That is, a way to 
write functions in a language of your choice and have them appear in Excel
as worksheet functions and macros.

xlOil is designed to have very low overheads when calling your own worksheet 
functions.

xlOil supports different languages via plugins. The languages currently 
supported are:

- C++
- Python
- SQL

In addition there is *xlOil_Utils* which contains some handy tools which Microsoft
never quite go around to adding.

You can use xlOil as an end-user of these plugins or you can use it to write
you own language bindings and contribute.

The latest documentation is here: https://xloil.readthedocs.io. 

You can build the documentation from a clone of the source with:

    cd docs
    make html


xlOil features
--------------

* Python
  - Very concise syntax to declare an Excel function
  - Optional type checking of function parameters
  - Supports keyword arguments
  - Choice of globally availble functions or code modules associated and limited to a single        workbook
  - Tight integration with numpy - very low overheads for array functions
  - Understands python tuples, lists, dictionarys and pandas dataframes
  - Async functions
  - Macro type functions which write to the sheet
  - Access to the Excel Application object
  - Hook Excel events
  - Pass any python object back to Excel and then back into any python function
  - Simple and quick add-in deployment
* C++
  - Safe and convenient wrappers around most things in the C-API
  - Concise syntax to declare Excel functions: registration is automatic
  - Deal with Excel variants, Ranges, Arrays and strings in a natural C++ fashion
  - Object cache allows returning opaque objects to Excel and passing them back to other functions
* SQL
  - Create tables from Excel ranges and arrays
  - Query and join them with the full sqlite3 SQL syntax
* Utils: very fast functions to:
  - Sort on multiple columns
  - Split and join strings
  - Make arrays from blocks

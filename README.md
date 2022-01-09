
xlOil
=====

xlOil provides framework for interacting with Excel in different programming
languages. It gives a way to write functions in a language and have them 
appear in Excel as worksheet functions or macros or have them control GUI elements. 
For example, the Python bindings can replace VBA in almost all use cases 
and even provide functionality not available from VBA.

xlOil is designed to have very low overheads when calling your worksheet 
functions.

xlOil supports different languages via plugins. The languages currently 
supported are:

- C++
- Python
- SQL

In addition there is *xlOil_Utils* which contains some handy tools which Microsoft
never quite got around to adding.

The latest stable documentation is here: https://xloil.readthedocs.io/en/stable.

xlOil features
--------------

* Python
  - Very concise syntax to declare an Excel function
  - Optional type checking of function parameters
  - Supports keyword arguments
  - Choice of globally declared functions or code modules limited to a single workbook just
    like VBA workbook-level functions
  - Tight integration with numpy - very low overheads for array functions
  - Understands python tuples, lists, dictionarys and pandas dataframes
  - Async functions
  - RTD functions and on-the-fly RTD server creation
  - Macro type functions which write to ranges on the sheet
  - Access to the Excel Application object and the full Excel object model
  - Hook Excel events
  - Pass any python object back to Excel and then use it as an argument for any python-based function
  - Simple and quick add-in deployment
  - Two-way connection to Jupyter notebooks: run worksheet functions in Jupyter and query variables
    in the jupyter kernel
  - Create Ribbons toolbars and Custom Task Panes
  - Return *matplotlib* plots and images from worksheet functions
* C++
  - Safe and convenient wrappers around most things in the C-API
  - Concise syntax to declare Excel functions: registration is automatic
  - Deal with Excel variants, Ranges, Arrays and strings in a natural C++ fashion
  - Object cache allows returning opaque objects to Excel and passing them back to other functions
  - Simplified RTD server creation
  - RTD-based background calculation
  - Create Ribbons toolbars and Custom Task Panes  
* SQL
  - Create tables from Excel ranges and arrays
  - Query and join them with the full sqlite3 SQL syntax
* Utils: very fast functions to:
  - Sort on multiple columns
  - Split and join strings
  - Make arrays from blocks


Supporting other languages
--------------------------

You can use xlOil as an end-user of these plugins or you can use it to write
you own language bindings (and ideally add them to the repo).
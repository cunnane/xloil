===========
xlOil
===========

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

In addition there is :any:`xlOil_Utils` which contains some handy tools which Microsoft
never quite go around to adding.

You can use xlOil as an end-user of these plugins or you can use it to write
you own language bindings and contribute.

**Start Here**  :any:`core-getting-started`

xlOil features
--------------

* Python
    - Very concise syntax to declare an Excel function
    - Optional type checking of function parameters
    - Supports keyword arguments
    - Choice of globally availble functions or code modules associated and limited to a single workbook
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

Why xlOil was created
---------------------

Programming with Excel has no ideal solution. You have a choice of poisons:

- VBA - great integration with Excel, but clunky syntax compared with
  other languages, few common libraries, limited IDE, testing and source 
  control very difficult, single-threaded, use not encouraged by MS.
- C-API - It's C and hence unsafe. The API is also old, has some quirks 
  and is missing many features of the other APIs. But it's the fastest
  interface.
- COM - More fully-featured but slower and strangely missing some features
  of the C-API.  It's C++ so still pretty unsafe. Has some unexpected behaviour
  and may fail to respond.  Requires COM binding support in your language 
  or the syntax is essentially unusuable.
- .Net API - actually sits on top of COM, the best modern interface
  but only for .Net languages and speed-limited by COM overheads, still missing 
  some C-API features.

xlOil tries to give you the C and COM APIs blended in a more friendly 
fashion and adds:

- Solution to the "how to register a worksheet function without a static DLL entry point" problem
- Object caching
- A framework for converting excel variant types to another language and back
- A convenient way of creating worksheet-scope functions
- A loader stub
- Goodwill to all men


Comparison between xlOil and other Excel Addin solutions
--------------------------------------------------------

Given the age of Excel, there are many other solutions for Excel add-in
development.

Most available packages are commercial (and quite expensive) so I 
cannot test relative performance with xlOil. They are generally
focussed on a single language rather than providing a framework
for language bindings as xlOil does.

The following, likely incomplete, list includes some of the most  
prominent Excel addin software. 

- Addin express: (commercial) fully-featured with nice GUI and the support
  seems good as well. Limited to .Net languages.
- ExcelDNA: (free) mature, well-regarded and widely used, covers 
  almost all Excel API features, but only for .Net languages.  I strongly
  recommend this software if you are using .Net languages.
- XLL Plus: (commercial) seems to be fully-featured with GUI wizards
  to help developers, but only for C++ and the most expensive sofware 
  here.
- PyXLL: (commercial) Python-only but supports the full range of Excel
  API features and some functionality to run on remote servers.
- XlWings: (mostly free) Python-only. More mature sofware, but considerably
  slower (2000x in my test case) than xlOil due to use of slower APIs.
  Can only create 'local' functions backed by VBA, so every Excel sheet needs
  to be be a macro sheet with special VBA redirects. This means it is not viable
  for addin deployment.  Supports Mac and Python 2.7 (but licence fee required 
  for this).
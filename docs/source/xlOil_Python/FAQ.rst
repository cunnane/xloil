=============================================
xlOil Python Questions (anticipated or asked)
=============================================

.. contents::
    :local:


ImportError: Typelib different than module
------------------------------------------

When using `comtypes` for COM support, then auto-generated modules can go out of sync, for example, if
you upgrade `comtypes`.  Fix this at a command prompt with 

.. highlight:: dosbatch

:: 

    C:\ > where clear_comtypes_cache.py
    C:\MyPythonDist\Scripts\clear_comtypes_cache.py
    C:\ > python C:\MyPythonDist\Scripts\clear_comtypes_cache.py -y

You'll probably need to restart Excel.


Intellisense / Function Context Help
------------------------------------

To activate pop-up function help, follow the instructions here: :any:`concepts-intellisense`.


Dynamically Resized Arrays
--------------------------

This is available in Office 365.  It would be possible to replicate this behaviour in older Excel 
versions however it is somewhat tricky, as the output ranges are not 'protected' as they are with 
array formulae or with the Office 365 support.  The code would need to:
 
    1. Hook the *AfterCalculate* event.
    2. Remember which functions output arrays in the current calc cycle and their calling cell.
    3. Remember which functions output arrays in the previous calc cycle and their calling cell.
    4. On *AfterCalculate*, loop through the functions in (2) writing their array to the worksheet.
    5. When writing the array, take care to clear any previous result from (3) but not to overwrite
       any other non-empty cells.
    6. Clear the ranges for functions in (3) but not in (2).
    7. Carefully handle the case where the output range for a function in (3) has been edited, for example
       it may now contain a function in (2)!

Unlike the built-in support in Office 365, the written arrays would be static data so, for example,
function dependency tracing would not work on them (except the top left entry).

This application failed to start because it could not find or load the Qt platform plugin "windows"
---------------------------------------------------------------------------------------------------
Sometimes Qt crashes with this error. It occurs when `QT_QPA_PLATFORM_PLUGIN_PATH` is not set correctly.
xlOil sets this when it lauches Qt but I think due to a race condition the message doesn't always get 
through.  Just restart Excel.
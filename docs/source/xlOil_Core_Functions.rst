=========================
xlOil Core Functions
=========================

This page describes the worksheet functions registered by the xlOil core, which should always 
be available when xlOil is running.

xloRef: adds a value or array to the cache
------------------------------------------

.. function:: xloRef(arrayOrRange)

    Adds the specified value or range or array to the object cache and returns a string 
    reference.

    See :any:`core-cached-objects`

xloVal: Retrives a value from the Excel data cache
--------------------------------------------------

.. function:: xloVal(cacheRefString)

    Given a reference string, returns a stored array or value. Cached values 
    are not saved with the workbook so will need to be recreated by forcing 
    a full recalc (press Ctrl-Alt-F9) when the workbook is opened

    See :any:`core-cached-objects`


xloHelp: returns help on an xlOil registered function
------------------------------------------------------

.. function:: xloHelp(functionName)

    Returns a two column array containing the help information passed to Excel's 
    function wizard when the function was registered.  The first row contains 
    the function name and the main help string. Subsequent rows contain argument
    names and their associated help.

    In the function wizard all help strings are limited to 256 characters. xloHelp
    does not have this limitation. 

xloLog: flushes the log file and returns its location
-----------------------------------------------------

.. function:: xloLog(ShowWindow=FALSE)

    The xlOil log is only flushed (written to file) occasionally or when an error 
    or warning occurs. Executing this function flushes the log and returns the 
    location of the log file (by default this is the same directory as the settings
    file).

    Setting the *ShowWindow* parameter causes xlOil's log window to appear. This 
    doesn't display the contents of the entire log file but only the most recent log 
    messages.


xloVersion: returns information on the xlOil version
----------------------------------------------------

.. function:: xloVersion()

    Returns a two-column array of version information.
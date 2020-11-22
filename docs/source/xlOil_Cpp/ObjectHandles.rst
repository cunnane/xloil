========================
xlOil C++ Object Handles
========================

It is often convenient to pass a pointer to an in-memory object from one function to another in a way
which generalises :any:`core-cached-objects`.  xlOil makes this very straightforward:

.. highlight:: c++

::

    XLO_FUNC_START(
        cacheOut(const ExcelObj& val)
    )
    {
        auto key = makeCached<int>(val.toInt());
        return returnValue(key);
    }
    XLO_FUNC_END(cacheOut);

    XLO_FUNC_START(
        cacheIn(const ExcelObj& cacheKey)
    )
    {
        auto* val = getCached<int>(cacheKey.asPString());
        return returnValue(val ? *val : 0);
    }
    XLO_FUNC_END(cacheIn);

The cache key or handle returned looks `<UnusualChar>[Book1]Sheet1!R1C1,A`. The leading character 
allows rapid rejection of invalid cache strings; it is unique per cached object type.

Cached objects do not persist when a spreadsheet is closed, but Excel does not know to recalculate
the cells which generate the cache handles when a spreadsheet is opened, so a full recalc must
be manually performed with Ctrl-Alt-F9.  To avoid this requirement, you can make the functions which
generate the handles RTD, but this carries some overhead.

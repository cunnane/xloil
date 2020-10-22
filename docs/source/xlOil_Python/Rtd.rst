======================
xlOil Python Async/Rtd
======================

Introduction
------------

In Excel, RTD (real time data) functions are able to return values independently of Excel's 
calculation cycle. The classic example of this is a stock ticker with live prices.  It is very 
easy to create an RTD function in xlOil_Python -- the following will give you a ticking clock:

::

    import xloil, datetime, asyncio

    @xloil.func
    async def pyClock():
        while True:
            await asyncio.sleep(2)
            yield datetime.datetime.now()

Note that you need calculation on automatic mode or you will not see the updates. Also note that
whatever parameter is passed to `sleep` the clock will not tick faster than 2 seconds - this is a 
limitation imposed by Excel.

In fact any `async` function is handled in xlOil as an RTD function by default.  It is possible to 
use Excel's native async support, but this has some drawbacks, discussed in :any:`concepts-rtd-async`

There is another advantage of RTD for xlOil: RTD functions can be 'local', i.e. called through a 
VBA stub associated with the workbook which avoids cluttering the global function namespace.


xlOil's RTD Interface
---------------------

Below we explain the details of the RTD mechanism for cases where finer-grained control is
required (such as in the `xloil_jupyter` module, :doc:`Jupyter`)

See the example in :doc:`Example`.

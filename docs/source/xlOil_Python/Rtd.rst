xlOil Python Async/Rtd
======================

Introduction
------------

In Excel, RTD functions are able to return values independently of Excel's calculation cycle.
The classic example of this is a stock ticker with live prices.  It is very easy to create 
an RTD function in xlOil_Python -- the following will give you a ticking clock:

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
use Excel's native async support, but this has some drawbacks, which we discuss below.


Comparison between Excel's async and RTD
----------------------------------------

Excel has supported RTD functions at least since Excel 2002.  In Excel 2010, Excel introduced 
native async functions.

RTD:

    * Pro: operates independently of the calc cycle - true background execution
    * Pro: provides notification when an RTD function call is changed or removed
    * Con: increased overhead compared to native async
    * Con: requires automatic calculation enabled (or repeated presses of F9 until calc is done)

Native async:

    * Pro: Lighter weight compared to RTD
    * Pro: works with manual calc mode
    * Con: tied to calc cycle, so any interruption cancels all async functions

The last con is particularly problematic for native async: *any* user interaction with Excel will
interrupt the calc, so whilst native async functions can run asynchronously with each other, they
cannot be used to perform background calculations.

There is another advantage of RTD for xlOil: RTD functions can be 'local', i.e. called through a 
VBA stub.


xlOil's RTD Interface
---------------------

Below we explain the details of the RTD mechanism for cases where finer-grained control is required
(such as in the `xloil_jupyter` module, :doc:`xlOil_Python_Jupyter`)

See the example in :doc:`xlOil_Python_Example`.

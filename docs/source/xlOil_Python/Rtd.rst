======================
xlOil Python Async/Rtd
======================

Introduction
------------

RTD (real time data) functions are able to return values independently of Excel's calculation
cycle. The classic example of this is a stock ticker with live prices.  It is easy to create
an RTD function in *xlOil_Python* -- the following gives a ticking clock:

::

    import xloil, datetime, asyncio

    @xloil.func
    async def pyClock():
        while True:
            yield datetime.datetime.now()
            await asyncio.sleep(2)
            

Note that calculation must be on automatic mode or you will not see the updates. 
Whatever parameter is passed to `sleep` the clock will not tick faster than 2 seconds - this due 
to the `RTD throttle interval <https://docs.microsoft.com/en-us/previous-versions/office/developer/office-xp/aa140060(v=office.10)>`_
Which can bed changed via `xlo.app().RTD.ThrottleInterval = <milliseconds>`, however 
reducing it below the default of 2000 may impair performance.  The change is global and
persists when Excel is restarted, so give some consideration to altering the value.

Any `async def` function is handled in xlOil as using RTD by default.  It is possible to 
use Excel's native async support, but this has some drawbacks, discussed in :any:`concepts-rtd-async`

There is another advantage of RTD: RTD functions can be 'local', i.e. called through a 
VBA stub associated with the workbook, which avoids cluttering the global function namespace.

Improving RTD performance (specifying topics)
---------------------------------------------

Registering an `async def` function as described above has a certain overhead: excel will 
call the function multiple times to fetch the result, so xlOil must store and compare all 
the function arguments to figure out if Excel wants the result of a previous calculation 
or to starta new calculation with new arguments.

If an RTD `topic`, i.e. a unique string identifier, is easy to determine we can take over
responsibility for generating it manually.

::

    # First create a new RTD COM server so the `topic` strings don't collide
    _rtdServer = xlo.RtdServer()
    
    @xloil.func
    def pyClock2(secs):

        async def fetch() -> dt.datetime:
            while True:
                await asyncio.sleep(secs)
                yield dt.datetime.now()
            
        return xloil.rtd.subscribe(_rtdServer, "Time:" + str(secs), fetch)

The `subscribe` call will look for an existing publisher, i.e. a clock with `secs` interval,
and return the value if one is found.  Otherwise it will run the coroutine and publish
the value.  Note the coroute specifies a return type: this is handled with a return converter
just like functions decorated with :any:`xloil.func`.

The instance of the :obj:`xloil.RtdServer` object creates and registers a COM class. When xlOil is
unloaded, the server will be unregistered and destroyed.  Since we created our own server the 
server, we can choose the convention for these topic strings (i.e. the unique publisher ID).

Note that we do not need to declare the outer function async, the `subscribe()` call notifies 
Excel that this function should be treated as RTD.

xlOil's RTD Interface
---------------------

If even finer-grained control of the RTD mechanism is required (such as in the `xloil_jupyter`
module, :doc:`Jupyter`), we can specify the publisher as described below.

We will follow the *UrlGetter* example in :doc:`Example`.  In this case we make the topics URLs. 
The RTD workflow is to first check if a given topic has a publisher using `peek()`. If not, 
we spin one up with :any:`xloil.RtdServer.start`. Then we :any:`xloil.RtdServer.subscribe` to 
the topic which tells xlOil to call Excel's RTD function.

:: 

    _rtdServer = xlo.RtdServer()

    @xlo.func
    def pyGetUrlLive(url):
        if _rtdServer.peek(url) is None:
            publisher = UrlGetter(url)
            _rtdServer.start(publisher)
        return _rtdServer.subscribe(url)


The publisher is the class which does the work. Its `connect()` method is called when a 
worksheet function calls `subscribe()` for its topic.  The publisher should then start
an async task to publish values.

If the worksheet function is subsequently changed or deleted, then `disconnect()` is called. 
When a publisher has no subscribers it should save CPU cycles by stopping its task.  A 
publisher should also stop when requested by the `stop()` method.

Apart from `connect()` the remaining methods are boilerplate at least for a simple publisher.
The boilerplate can be avoided by use of the `RtdSimplePublisher` class, then only the
`run()` method in the below requires definition. 

::

    class UrlGetter(xlo.RtdPublisher):

        def __init__(self, url):
            # You *must* call this ctor explicitly or the python binding library will crash
            super().__init__()  
            self._url = url
            self._task = None
           
        def connect(self, num_subscribers):
            if self.done():
                async def run():
                    try:
                        while True:
                            data = await getUrlAsync(self._url);
                            _rtdServer.publish(self._url, data)
                            await asyncio.sleep(4)                     
                    except Exception as e:
                        _rtdServer.publish(self._url, e)
                        
                self._task = xlo.get_event_loop().create_task(run())
                
        def disconnect(self, num_subscribers):
            if num_subscribers == 0:
                self.stop()
                # Returning True schedules the publisher for destruction
                return True 
                
        def stop(self):
            if self._task is not None: 
                self._task.cancel()
        
        def done(self):
            return self._task is None or self._task.done()
            
        def topic(self):
            return self._url

The final task, left as an exercise, is to write `getUrlAsync()`: an async function which 
fetches a URL.  It is straightforward with the `aiohttp` library.

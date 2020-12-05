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

Note that you need calculation on automatic mode or you will not see the updates. Also,
whatever parameter is passed to `sleep` the clock will not tick faster than 2 seconds - this due 
to the `RTD throttle interval <https://docs.microsoft.com/en-us/previous-versions/office/developer/office-xp/aa140060(v=office.10)>`_
Which can bed changed via `xlo.app().RTD.ThrottleInterval = <milliseconds>`, however 
reducing it below the default of 2000 may impair performance.  Since the change is global and
persists when Excel is restarted, give some consideration to altering the value.

Any `async` function is handled in xlOil as an RTD function by default.  It is possible to 
use Excel's native async support, but this has some drawbacks, discussed in :any:`concepts-rtd-async`

There is another advantage of RTD for xlOil: RTD functions can be 'local', i.e. called through a 
VBA stub associated with the workbook, which avoids cluttering the global function namespace.


xlOil's RTD Interface
---------------------

Below we explain the details of the RTD mechanism for cases where finer-grained control is
required (such as in the `xloil_jupyter` module, :doc:`Jupyter`).

We will follow the *UrlGetter* example in :doc:`Example`.

First create an instance of the `RtdServer` object. This creates and registers a COM class.
When xlOil is unloaded, the server will be unregistered and destroyed.

The server mantains a dict of topic string to `IRtdPublisher` tasks. Since we control the 
server, we choose the convention for these topic strings: in this case we make them URLs. 
The RTD workflow is to first check if a given topic has a publisher using `peek()`. If not, 
we spin one up with `start()`. Then we `subscribe()` to the topic which tells xlOil to called
Excel's RTD function.

:: 

    _rtdServer = xlo.RtdServer()

    @xlo.func
    def pyGetUrlLive(url):
        if _rtdServer.peek(url) is None:
            publisher = UrlGetter(url)
            _rtdServer.start(publisher)
        return _rtdServer.subscribe(url)

Note that we do not need to declare the function async, the `subscribe()` call notifies Excel
that this function should be treated as RTD.

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

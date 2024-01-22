import xloil as xlo
import sys
import datetime as dt
import asyncio
import os 
import numpy as np
 
#------------------
# Async functions
#------------------
#
# Using asyncio's async keyword declares an async function in Excel.
# This means control is passed back to Excel before the function 
# returns.  Python is single-threaded so no other python-based functions
# can run whilst waiting for the async return. However, the await keyword 
# can pass control between running async functions.
#
# There are two flavours of async function: RTD and native. The XLL interface
# contains async support but any interaction with Excel will cancel all native async 
# functions: they are only asynchronous with each other, not with the user interface.
# This is fairly unexpected and generally undesirable, so xlOil has an implementation of 
# async which works in the expected way using RTD at the expense of more overhead.
#
@xlo.func
async def pyTestAsyncRtd(x, time:int):
    await asyncio.sleep(time)
    return x
    
#
# Native async functions cannot be declared local as VBA does not support this. We do
# not actually need to specify local=False, as xloil will automatically set this.
# 
@xlo.func(rtd=False, local=False)
async def pyTestAsync(x, time:int):
    await asyncio.sleep(time)
    return x

@xlo.func
async def pyTestAsyncGen(secs):
    while True:
        await asyncio.sleep(secs)
        yield dt.datetime.now()

@xlo.func(local=False, threaded=True)
async def pyTestAsyncThreaded(secs):
    while True:
        await asyncio.sleep(secs)
        yield dt.datetime.now()
        
@xlo.func
async def pyRtdArray(values):
    return np.sum(values)

#
# Retrieve the address of calling cell (assuming we are called from a sheet)
# Note we don't need macro permissions to do this.
#  
@xlo.func
async def pyTestCaller():
    return xlo.Caller().address()
    
#---------------------------------
# Calling Excel built-in functions
#---------------------------------
#
# This can be done asynchronously as Excel built-ins can only be called on the 
# main thread.
#
@xlo.func
async def pyTestExcelCallAsync(x, y, z):
    return await xlo.call_async("sum", x, y, z)

@xlo.func   
def pyTestExcelCall(func, arg1:xlo.AllowRange=None, arg2:xlo.AllowRange=None, arg3:xlo.AllowRange=None):
    # We pop the trailing missing args so the called function 
    # receives the correct number of arguments. None is converted
    # to Missing when calling Excel built-ins
    args = [arg1, arg2, arg3]
    while args[-1] is None:
        args.pop()
    return xlo.call(func, *args)
    
@xlo.func   
def pyTestAppRun(func, arg1:xlo.AllowRange=None, arg2:xlo.AllowRange=None, arg3:xlo.AllowRange=None):
    return xlo.run(func, arg1, arg2, arg3)
    
#---------------------------------
# RTD functions and the RTD server
#---------------------------------
#
# Registering an `async def` function has a certain overhead:
# Excel will call your function multiple times to fetch the result
# So xlOil must store and compare all the function arguments to figure
# out if Excel wants the result of a previous calculation or to start
# a new calculation with new arguments.
# 
# If the RTD `topic`, i.e. the unique identifier, is easy to determine
# we can take over responsibility for generating it ourselves.
#  
# First create a new RTD COM server so the `topic` strings don't collide
_rtdServer = xlo.RtdServer()
  
@xlo.func
def pyTestRtdManual(secs):

    # This coroutine will be run if  we don't already have a 
    # publisher for the specified number of seconds.
    async def fetch() -> dt.datetime:
        while True:
            await asyncio.sleep(secs)
            yield dt.datetime.now()
        
    return xlo.rtd.subscribe(_rtdServer, "Time:" + str(secs), fetch)

#
# Now try a slightly more practical usage of RTD async: fetching URLs.  
# (We need the aiohttp package for this).  Here we use the RTD machinery
# in full-manual mode, defining the publishing object explicitly. This
# is not necessary, it's just illustrative.
#
try:
    import aiohttp
    import ssl

    # This is the implementation: it pulls the URL and returns the response as text
    async def _getUrlImpl(url):
        async with aiohttp.ClientSession() as session:
            async with session.get(url, ssl=ssl.SSLContext()) as response:
               return await response.text() 
        
    
    #
    # We declare an async gen function which calls the implementation either once,
    # or at regular intervals
    #
    @xlo.func(local=False, rtd=True)
    async def pyGetUrl(url, seconds=0):
        yield await _getUrlImpl(url)
        while seconds > 0:
            await asyncio.sleep(seconds)
            yield await _getUrlImpl(url)
             

    #
    # Below we show how to write the above function in "long form" with
    # explicit connections to the RtdManager. In our implementation below
    # we repeatedly poll the URL every 4 seconds, This is just an example 
    # to show how to use the full RTD functionality: in general it is 
    # better to let xlOil handle things and use an async generator.
    # 

    # 
    # RTD servers use a publisher/subscriber model with the 'topic' as the
    # key. The publisher below is linked to a single topic string, which is the 
    # url to be fetched. 
    # 
    # We have designed the publisher to do nothing on construction. When it detects
    # a subscriber, it creates a publishing task on xlOil's asyncio loop (which runs
    # in a background thread). When there are no more subscriber, it cancels this task.
    # If the task was very slow to return, we could have opted to start it in the constructor  
    # and kept it running permanently, regardless of subscribers.
    # 
    class UrlGetter(xlo.RtdPublisher):

        def __init__(self, url):
            super().__init__()  # You *must* call this explicitly or the python binding library will crash
            self._url = url
            self._task = None
           
        def connect(self, num_subscribers):
        
            if self.done():
            
                async def run():
                    try:
                        while True:
                            data = await _getUrlImpl(self._url);
                            _rtdServer.publish(self._url, data)
                            await asyncio.sleep(4)                     
                    except Exception as e:
                        _rtdServer.publish(self._url, e)
                        
                self._task = xlo.get_event_loop().create_task(run())
                
        def disconnect(self, num_subscribers):
            if num_subscribers == 0:
                self.stop()
                return True # This publisher is no longer required: schedule it for destruction
                
        def stop(self):
            if self._task is not None: 
                self._task.cancel()
        
        def done(self):
            return self._task is None or self._task.done()
            
        def topic(self):
            return self._url
    
    
    @xlo.func(local=False)  
    def pyGetUrlLive(url):
        # We 'peek' into the RTD manager to see if there is already a publisher for 
        # our topic. If not we create one, then issue the subscribe request, which 
        # registers the calling cell with Excel as an RTD cell.
        if _rtdServer.peek(url) is None:
            publisher = UrlGetter(url)
            _rtdServer.start(publisher)
        return _rtdServer.subscribe(url)       
    
except ImportError:
    @xlo.func(local=False)
    def pyGetUrl(url):
        return "You need to install aiohttp" 

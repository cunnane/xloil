import xloil as xlo
from xloil.pandas import PDFrame
import datetime as dt
import asyncio


#
# Functions are registered by decorating them with xloil.func.  The function
# doc-string will be displayed in Excel's function wizard
#
@xlo.func
def pySum(x, y, z):
    '''Adds up numbers'''
    return x + y + z

#
# If argument types or function return types are specified using 'typing' 
# annotations, xloil # will attempt to convert Excel's value to the specified
# type and  will throw if it can't. 
# 
# Argument defaults using the normal python syntax are respected.
#
@xlo.func
def pySumNums(x: float, y: float, a: int = 2, b: int = 3) -> float:
	return x * a + y * b
    
#
# The registered function name can be overriden as can the doc-string.
# The 'group' argument specifes a category of functions in Excel's 
# function wizard
#
@xlo.func(
    name='pyRoundTrip', 
    group='UselessFuncs', 
    help='returns its argument',
    args={'x': 'the argument'}
    )
def pyTest1(x):
    '''
    Long description, too big for function wizard, which is actually limited
    to 255 chars, presumably because, despite it being quite central to Excel
    the function wizard hasn't been improved in 20 years.... The icons on the
    other hand...
    '''
    return x

#
# Ranges (e.g. A1:B2) passed as arguments are converted to numpy arrays
# The default numpy dtype is object, but it's more performant to specify
# a dtype if you can.  xlOil will raise an error if it cannot make the
# conversion.
#
@xlo.func(args={'x': "2-dim array to return"})
def pyTestArr2d(x: xlo.Array(float)):
	return x

#
# By default, ranges are trimmed to the last non-empty row and column.
# Non empty is any value which is not #N/A or a zero length string
# or an empty cell.  This default is desiable as it allows input from
# functions which return a variable length array (which Excel pads with 
# #N/A when writing to the sheet) or variable length user input.  This
# behaviour can be disabled as shown below.
#
# Note you cannot use keyword args in [], see PEP472
#
@xlo.func
def pyTestArrNoTrim(x: xlo.Array(object, trim=False)):
	return x


# 
# This func uses the explicit `args` specifier with xlo.Arg. This overrides any
# auto detection of the argument type or default by xlOil.
# 
@xlo.func(args=[ 
        xlo.Arg("multiple", typeof=float, help="value to multiply array by", default=1)
    ])
def pyTestArr1d(x: xlo.Array(float, dims=1), multiple):
	return x * multiple


#------------------
# The Object Cache
#------------------
#
# If you attempt to return a non-convertible object to Excel, xlOil
# will store it in a cache an instead return a reference string based
# on the currently calculating cell. 
# 
# To use this returned value in another function, do not specify an argument
# type. xlOil will check if the provided argument is a reference to a cache 
# objects and, if so, fetch it and pass it to the function.
#

class CustomObj:
    def __init__(self):
        self.greeting = 'Hello world'
    
@xlo.func
def pyTestCache(cachedObj=None):
    """
    Returns a cache reference to a greeting object if no argument is provided.
    If a greeting object is given, returns the greeting as text.
    """
    if type(cachedObj) is CustomObj:
        return cachedObj.greeting
    return CustomObj()
   
#------------------
# Dates
#------------------
#
# xlOil can convert Excel values to dates but:
#   1) You must specify the argument type as date or datetime. Excel
#      stores dates as numbers so xlOil cannot know when a date
#      conversion is required (because it uses the XLL interface)
#   2) Excel dates cannot contain timezone information
#   3) Excel dates cannot be before 1 Jan 1900 or after December 31, 9999
# We don't specify a datetime return type
# 
@xlo.func
def pyTestDate(x: dt.datetime) -> dt.datetime:
    return x + dt.timedelta(days=1)
 

#
# Keyword args are supported by passing a two-column array of (string, value)
# This function also tests the dict return conversion (without specifying the
# return as dict, the iterable converter would be used resulting in output of
# only the keys)
#
@xlo.func
def pyTestKwargs(**kwargs) -> dict:
    return kwargs

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


#---------------------------------
# RTD functions and the RTD server
#---------------------------------
#
# We try a slightly more practical usage of RTD async: fetching URLs.  
# (We need the aiohttp package for this)
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
    
    # We create a new RTD COM server
    _rtdServer = xlo.RtdServer()

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
                        _rtdServer.publish(self._url, str(e))
                        
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

#------------------
# Other handy bits
#------------------
#
# If an iterable object is returned, xlOil attempts to convert it to an
# array, with each element as a column. So a 1d iterator gives a 1d array
# and a iterator of iterator gives a 2d array.
# 
# If you want an iterable object to be placed in the cache use 
# `return xlo.to_cache(obj)`
#
@xlo.func
def pyTestIter(size:int, dims:int):
    if dims == 1:
        return [1] * size
    elif dims == 2:
        return [[1] * size] * size
    else:
        return [] 

#
# Declaring a function as a macro allows use of the Excel.Application object
# accessible via `xlo.app()`. The available methods and properties are described
# in the microsoft documentation
#
@xlo.func(macro=True)
def pyTestCom():
    app = xlo.app()
    return app.ProductCode

#
# The special xlo.Range annotation allows the function to receive range arguments
# as an ExcelRange object. This allows extraction of part of the data without making a 
# copy of the entire range as an array.
#
@xlo.func(macro=True)
def pyTestRange(r: xlo.Range):
    r2 = r.cells(0, 0).value
    return r.cells(1, 1).address()

#
# Displays python's sys.path. Useful for debugging some module loads
# 
import sys        
@xlo.func(local=False)
def pysyspath():
    return sys.path

#
# Displays python's sys.modules. Useful for debugging some module loads
# 
@xlo.func(local=False)
def pysysmodules():
    return sys.modules
    
#
# Threads
# 
import numpy as np
import ctypes

@xlo.func(local=False, threaded=True)
def pyThreadTest(x: float, y: float, a: int, b: int, u:int, v:int) -> int:
    # Do something numpy intensive to allow thread switching
    np.sum(np.ones((a, b)) * x ** (np.ones((u, v)) / y))
    # Return the thread ID to prove the functions were executed on different threads
    return ctypes.windll.kernel32.GetCurrentThreadId(None)
    
#--------------------------------
# Custom argument type converters
#---------------------------------
#
# The `converter` decorator tells xlOil that the following function or 
# class is a type converter. A type converter creates a python object
# from a given bool, float, int, str, ExcelArray or ExcelRange.
#
# The converter can be applied to an argument using the usual annotation
# syntax, or using the `args` argument to xlo.func().
# 
@xlo.converter()
def arg_doubler(x):
    if isinstance(x, xlo.ExcelArray):
        x = x.to_numpy()
    return 2 * x

@xlo.func
def pyTestCustomConv(x: arg_doubler):
    return x

#-------------------
# Pandas Dataframes
#-------------------
#

try:
    import pandas as pd

    #
    # xlo.PDFrame converts a block to a pandas DataFrame. The block should be
    # formatted as a table with data in columns and a row of column headings
    # if the headings parameter is set. We have to explicitly add the return
    # value to the cache or it will be expanded to the sheet
    #
    @xlo.func(args={'df': "Data to be read as a pandas dataframe"})
    def pyTestDFrame(df: PDFrame(headings=True)):
        return xlo.cache.add(df)

    #
    # We can tell xlo.PDFrame to set the datafram index to a specified column 
    # name. If you want the index column name to be dynamic, you'll need to
    # drop the index param and call DataFrame.set_index yourself.
    #
    @xlo.func
    def pyTestDFrameIndex(df: PDFrame(headings=True, index="Time")):
        return xlo.cache(df)

    #
    # This function tests that we can fetch data from the frames created by the
    # previous functions
    #
    @xlo.func
    def pyTestFrameFetch(df, index=None, col_name=None):
        if index is not None:
            if col_name is not None:
                return df.loc[index, col_name]
            else:
                return df.loc[index].values
        elif col_name is not None:
            return df[col_name]
        else:
            return df
        
except ImportError:
    pass

#-----------------
# Event handling 
#-----------------
#
# We setup some simple event handlers and demonstrate some more
# use of of the app() object and using Range. 
#
# Currently event handlers are global, so for workbook local modules
# such as this one, we check compare the the workbook name to the global
# `_xloil_workbook` which is set by xlOil when the module is imported
#
@xlo.func
def getLinkedWbName():
    return _xloil_workbook
    
def event_writeTimeToA1():
    if xlo.app().ActiveWorkbook.Name != _xloil_workbook:
        return
        
    sheet_name = xlo.app().ActiveSheet.Name
    time = str(dt.datetime.now())
    range = xlo.Range("A1")
    range.value = f"Calc on {sheet_name} finished at: {time}"

#
# This handler is for the WorkbookBeforePrint event. If the `cancel` parameter
# is set to True, the print is cancelled. Since python does not support changing
# bool function arguments directly (i.e. reference parameters), we must use the
# syntax `cancel.value = True`
#
def event_stopPrinting(wbName, cancel):
    if wbName != _xloil_workbook:
        return
    xlo.Range("B1").value = "Cancelled print for: " + wbName
    cancel.value = True

#
# Link the above handlers to events. To unlink them, use `-=`. Note that
# xlOil only holds weak references to the event handler functions, so they must
# be made module scope variables to stay alive, i.e. the following will not work:
#
#       xlo.event.AfterCalculate += lambda x: <do something>
# 
# Rather write:
# 
#       _handler = lambda x: <do something>
#       xlo.event.AfterCalculate += _handler
#
# The advantage of the weak reference is that the handler is automatically unlinked
# when the containing module is unloaded, so there is no need to explictly do `-=`
# in the `_xloil_unload` function.
#

xlo.event.AfterCalculate += event_writeTimeToA1
xlo.event.WorkbookBeforePrint += event_stopPrinting

#----------------------
# GUI: Creating Ribbons
#----------------------
#
# xlOil is able to create a ribbon entry for a workbook which is automatically 
# removed when the workbook is closed and the associated workbook module is 
# unloaded.  To create the ribbon XML use an editor such as Office RibbonX Editor:
# https://github.com/fernandreu/office-ribbonx-editor
#
def press1(ctrl):
    xlo.log("1 Pressed")
    pass
    
_ribbon = xlo.create_ribbon(r'''
    <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
        <ribbon>
            <tabs>
                <tab id="customTab" label="xloPyTest" insertAfterMso="TabHome">
                    <group idMso="GroupClipboard" />
                    <group idMso="GroupFont" />
                    <group id="customGroup" label="MyButtons">
                        <button id="pyButt1" label="Button1" size="large" onAction="press1" imageMso="Bold" />
                    </group>
                    <group idMso="GroupEnterDataAlignment" />
                    <group idMso="GroupEnterDataNumber" />
                    <group idMso="GroupQuickFormatting" />
                </tab>
            </tabs>
        </ribbon>
    </customUI>
    ''', 
    mapper={
        'press1': press1
    })
    

#
# xlOil will attempt to call a function with this name when the module is unloaded,
# for example, because the linked workbook is closed. xlOil explictly clears the 
# module's __dict__ before unload, so any globals, like _ribbon above will be 
# deleted.
#
def _xloil_unload():
    pass


import xloil.debug
xloil.debug.exception_debug('pdb')

@xlo.func
def pyTestDebug(x):
    return (2 * x) ^ (x + 1)

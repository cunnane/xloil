import xloil as xlo
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
@xlo.func(name='pyRoundTrip', group='Useless', help='returns its argument')
def pyTest1(x):
    '''Long description, too big for function wizard'''
    return x

#
# Ranges (e.g. A1:B2) passed as arguments are converted to numpy arrays
# The default numpy dtype is object, but it's more performant to specify
# a dtype if you can.  xlOil will raise an error if it cannot make the
# conversion
#
@xlo.func
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
# This func uses the explicit `arg` specifier to allow it to give
# per-argument help strings
# 
@xlo.func
@xlo.arg("multiple", typeof=float, help="value to multiply array by")
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
#
@xlo.func
def pyTestDate(x: dt.datetime):
    return x + dt.timedelta(days=1)
 

# 
# Keyword args are supported by passing a two-column array of (string, value)
#
@xlo.func
def pyTestKwargs(argName, **kwargs):
    return kwargs[argName]

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
# Note that any interaction with Excel will cancel all async functions:
# they are only asynchronous with each other, not with the user interface.
#
@xlo.func(local=False)
async def pyTestAsync(x, time:int):
    await asyncio.sleep(time)
    return x

@xlo.func(thread_safe=True, local=False)
async def pyTestAsyncThread(x, time:int):
    await asyncio.sleep(time)
    return x

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
# accessible via `xlo.app()`
#
@xlo.func(macro=True)
def pyTestCom():
    app = xlo.app()
    return app.ProductCode

#
# The special xlo.AllowRange annotation allows the function to receive ranges
# as an ExcelRange object. This allows manipulation without making a copy of
# the data in the range.
#
@xlo.func(macro=True)
def pyTestRange(r: xlo.AllowRange):
    r2 = r.cells(1, 1).value
    return r.cells(2, 2).address()

#--------------------------------
# Custom argument type converters
#---------------------------------
#
# The `converter` decorator tells xlOil that the following function or 
# class is a type converter. A type converter creates a python object
# from a given bool, float, int, str, ExcelArray or ExcelRange.
#
# The converter can be applied to an argument using the usual annotation
# syntax, or using the `arg` specfier.
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
# xlo.PDFrame converts a block to a pandas DataFrame. The block should be
# formatted as a table with data in columns and a row of column headings
# if the headings parameter is set
#
@xlo.func
def pyTestDFrame(df: xlo.PDFrame(headings=True)):
    return xlo.cache.add(df)

#
# We can tell xlo.PDFrame to set the datafram index to a specified column 
# name. If you want the index column name to be dynamic, you'll need to
# drop the index param and call DataFrame.set_index yourself.
#
@xlo.func
def pyTestDFrameIndex(df: xlo.PDFrame(headings=True, index="Time")):
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
    else:
        return df[col_name]

#-----------------
# Event handling 
#-----------------
#
# We setup some simple event handlers and demonstrate some more
# use of of the app() object and using Range. 
#
# Currently event handlers are global, so for workbook local modules
# such as this one, we check compare the the workbook name to the global
# `_xl_this_workbook` which is set by xlOil when the module is imported
#
@xlo.func
def getLinkedWbName():
    return _xl_this_workbook
    
def event_writeTimeToA1():
    if xlo.app().ActiveWorkbook.Name != _xl_this_workbook:
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
    if wbName != _xl_this_workbook:
        return
    xlo.Range("B1").value = "Cancelled print for: " + wbName
    cancel.value = True

#
# Link the above handlers to the events. To unlink them, use `-=`.
#
xlo.event.AfterCalculate += event_writeTimeToA1
xlo.event.WorkbookBeforePrint += event_stopPrinting

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


@xlo.func
@xlo.arg("multiple", typeof=float, help="value to multiply array by")
def pyTestArr1d(x: xlo.Array(float, dims=1), multiple):
	return x * multiple

class CustomObj:
    def __init__(self):
        self.greeting = 'Hello world'

#
# If you attempt to return a non-convertable object to Excel, xloil
# will store it in a cache an instead return a reference string based
# on the currently calculating cell. If you do not specify a function
# argument type, xloil will check if any provided arguments are references
# to cache objects and, if so, fetch them.
#
@xlo.func
def pyTestCache(cachedObj=None):
    if type(cachedObj) is CustomObj:
        return cachedObj.greeting
    return CustomObj()
   

#
# xlOil can convert Excel values to dates but:
#   1) You must specify the argument type as date or datetime. Excel
#      stores dates as numbers so xlOil cannot know when a date
#      conversion is required
#   2) Excel dates cannot contain timezone information
#   3) Excel dates cannot be before 1900 or after ???
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

#
# Using asyncio's async keyword declares an async function in Excel.
# This means control is passed back to Excel before the function 
# returns.  Python is single-threaded so no other python-based functions
# can run whilst waiting for the async return unless they are also 
# declared async and the await keyword is used.
#
@xlo.func
async def pyTestAsync(x, time:int):
    await asyncio.sleep(time)
    return x

@xlo.func(thread_safe=True)
async def pyTestAsyncThread(x, time:int):
    await asyncio.sleep(time)
    return x
    
@xlo.func
def pyTestIter(size:int, dims:int):
    if dims == 1:
        return [1] * size
    elif dims == 2:
        return [[1] * size] * size
    else:
        return [] 

@xlo.func(macro=True)
def pyTestCom():
    app = xlo.app()
    return app.ProductCode


@xlo.func(macro=True)
def pyTestRange(r: xlo.AllowRange):
    r2 = r.cell(1, 1).value
    return r.cell(1, 1).address()

@xlo.converter()
def arg_doubler(x):
    if isinstance(x, xlo.ExcelArray):
        x = x.to_numpy()
    return 2 * x

@xlo.func
def pyTestCustomConv(x: arg_doubler):
    return x


@xlo.converter()
def testcon(x):
    if isinstance(x, xlo.ExcelArray):
        return x.to_numpy(dims=1).astype(object)
    return "#NAA"

@xlo.func
def pyTestCon1(x: testcon):
    return x


@xlo.func
def pyTestDFrame(df: PDFrame2(headings=True), col_name:str):
    return df[col_name]
import xloil as xlo
import sys
import datetime as dt
import asyncio
import os 
import ctypes

@xlo.func(local=False, threaded=True)
async def pyRtdThreadedClock(secs):
    while True:
        yield dt.datetime.now().second + ctypes.windll.kernel32.GetCurrentThreadId(None)
        await asyncio.sleep(secs)
        

@xlo.func(local=False, threaded=False)
async def pyRtdClock(secs):
    while True:
        yield dt.datetime.now().second
        await asyncio.sleep(secs)
        
        
@xlo.func(local=False, threaded=False)
async def pyRtdArrayGen(secs):
    while True:
        now = dt.datetime.now()
        yield [now.hour, now.minute, now.second]
        await asyncio.sleep(secs)
        
@xlo.func
async def pyRtdArray(values: xlo.Array(float,dims=1)):
    return [str(x) for x in values.tolist()]
    

import time

try:
    import xloil as xlo

    @xlo.func
    def xoTime(x):
        return time.clock()
except:
    pass

try:
    import xlwings as xw
    @xw.func
    def xwTime(name):
        return time.clock()

except:
    pass



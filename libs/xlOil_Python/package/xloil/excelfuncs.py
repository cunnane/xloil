import xloil as xlo

@xlo.func
def xloPyLoad(ModuleName:str):
    return str(xlo.scan_module(ModuleName))
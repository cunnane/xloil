import xloil as xlo

@xlo.func(help="Imports the specifed python module and scans it for xloil functions")
def xloPyLoad(ModuleName:str):
    return str(xlo.scan_module(ModuleName))
import xloil as xlo
import sys

@xlo.func(macro=True, 
          help="Imports the specifed python module and scans it for xloil functions")
def xloPyLoad(ModuleName:str = ""):

    workbook_name = None

    if len(ModuleName) == 0:
        import os
        caller = xlo.Caller()
        ModuleName = 'xloil_wb_' + os.path.splitext(caller.workbook)[0]
        if ModuleName not in sys.modules:
            full_path = xlo.app().Workbooks(caller.workbook).FullName
            ModuleName = os.path.splitext(full_path)[0] + ".py"
            workbook_name = caller.workbook
    else:
        # Little bit of a hacky way of getting to the workbook module
        wb_name = 'xloil_wb_' + ModuleName
        if wb_name in sys.modules:
            ModuleName = wb_name

    return str(xlo.scan_module(ModuleName, workbook_name))

@xlo.func(macro=True, 
          help="Imports the specifed python module and scans it for xloil functions")

def xloPyDebug(Debugger:str = ""):
    import xloil.debug
    xloil.debug.exception_debug(Debugger if len(Debugger) > 0 else None)
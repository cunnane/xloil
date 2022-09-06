import xloil as xlo
import sys
import importlib

@xlo.func(macro=True)
def xloPyLoad(ModuleName:str = ""):

    """Imports the specifed python module and scans it for xloil functions"""

    workbook_name = None

    from xloil.importer import _import_file_and_scan

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

    module = importlib.reload(sys.modules[ModuleName]) \
        if ModuleName in sys.modules else \
            _import_file_and_scan(ModuleName, workbook_name=workbook_name)
    
    return str(module)


@xlo.func(macro=True,
          args={
              'Debugger': "Choose from 'pdb', 'vscode', or empty string to disable"
              })
def xloPyDebug(Debugger:str = ""):
    """
    Sets the user-code exception debugger. 
    Pdb opens in a new window.
    """

    if xlo.in_wizard():
        return

    import xloil.debug
    if len(Debugger) > 0:
        xloil.debug.use_debugger(Debugger)
        return Debugger
    else:
        xloil.debug.use_debugger(None)
        return "OFF"

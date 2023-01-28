import xloil as xlo
import sys
import importlib
import os

@xlo.func(
    help="Loads functions from a module and registers them in Excel. The functions do not have to be decorated. "
         "Provides an analogue of `from X import Y as Z`",
    args={'ModuleName': 'A module name or a full path name to a target py file. If empty, the workbook module '
                        'with the same name as the calling workbook is (re)loaded.',
          'From': 'If missing, imports the specified module as normal. If a value or array, register only '
                  'the specified object names.  If "*", all objects are registered.',
          'To': 'Optionally specifies the Excel function names to register in the same order as `From`. '
                'Should have the same length as `From`' 
        })
async def xloImport(ModuleName:str, From=None, As=None):

    from xloil.importer import _import_file

    caller = xlo.Caller()
    workbook_name = os.path.splitext(caller.workbook)[0]

    # If no module name provided, try to load (or reload) the corresponding 
    # workbook module
    if len(ModuleName) == 0:
        ModuleName = f'xloil_wb_{workbook_name}'
        # If module does not exist, replace module name with full path 
        if ModuleName in sys.modules:
            importlib.reload(sys.modules[ModuleName])
        else:
            full_path = xlo.app().Workbooks(caller.workbook).FullName
            module_path = os.path.splitext(full_path)[0] + ".py"
            _import_file(module_path, workbook_name=workbook_name)
    elif From is None:
        module = importlib.reload(sys.modules[ModuleName]) \
                    if ModuleName in sys.modules else \
                        _import_file(ModuleName, workbook_name=workbook_name)
    else:
        xlo.import_functions(ModuleName, From, As, workbook_name=workbook_name)

    return f"Loaded {ModuleName}"

def _xlo_attr_helper(Object, Name:str, *Args, **Kwargs):
    attr = getattr(Object, Name)

    import inspect
    if inspect.ismethod(attr) or inspect.isfunction(attr):
        return attr(*Args, **Kwargs)
    else:
        return attr

@xlo.func(
    help="Returns the named attribute value, or the result of calling it if possible. "
         "ie, `object.attr` or `object.attr(*args, *kwargs)`",
    args={'Object': 'The target object',
          'Name': 'The name of the attribute to be returned.  The attribute can be a '
                  'bound method, member, property, method, function or class',
          'Args': 'If the attribute is callable, it will be called using these positional arguments',
          'Kwargs': 'If the attribute is callable, it will be called using these keyword arguments'
    })
def xloAttr(Object, Name:str, *Args, **Kwargs):
    return _xlo_attr_helper(Object, Name, *Args, **Kwargs)

@xlo.func(
    help="Returns the value of named attribute or the result of calling the attribute as a Cache object"
         "(cf `xloAttr`). This function is useful to stop default conversion to Excel values, for example "
         "when chaining xloAttr calls",
    args={'Object': 'The target object',
          'Name': 'The name of the attribute to be returned.  The attribute can be a '
                  'bound method, member, property, method, function or class',
          'Args': 'If the attribute is callable, it will be called using these positional arguments',
          'Kwargs': 'If the attribute is callable, it will be called using these keyword arguments'
    })
def xloAttrObj(Object, Name:str, *Args, **Kwargs) -> xlo.Cache:
    return _xlo_attr_helper(Object, Name, *Args, **Kwargs)

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

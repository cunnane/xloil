"""
Importing this module hooks the import and reload functions
"""

import importlib
import importlib.util
import importlib.abc
import builtins
import sys
import os
import inspect

from .register import scan_module, _clear_pending_registrations
from ._core import StatusBar
from ._common import log_except

_module_addin_map = dict()
_linked_workbooks = dict()

class _ModuleFinder(importlib.abc.MetaPathFinder):

    """
    Allows importing a module from a path specified in path_map
    without needing to add it to sys.paths - essentially a private
    set of import paths, indexed by module name
    """

    path_map = dict()

    def find_spec(self, fullname, path, target=None):
        path = self.path_map.get(fullname, None)
        if path is None:
            return None
        return importlib.util.spec_from_file_location(fullname, self.path_map[fullname])

    def find_module(self, fullname, path):
        return None

# We maintain a _ModuleFinder on sys.meta_path to catch any reloads of our off-path
# imported modules loaded by _import_file
_module_finder = _ModuleFinder()
sys.meta_path.append(_module_finder)


def linked_workbook(mod=None):
    """
        Returns the full path of the workbook linked to the specified module
        or None if the module was not loaded with an associated workbook.
        If no module is specified, the calling module is used.
    """
    if mod is None:
        # Get caller
        frame = inspect.stack()[1]
        mod = inspect.getmodule(frame[0])
    return _linked_workbooks.get(mod.__name__, None)


def source_addin(mod=None):
    if mod is None:
        # Get top-level caller
        frame = inspect.stack()[-1]
        mod = inspect.getmodule(frame[0])
    return _module_addin_map.get(mod.__name__, None)


def _import_scan(what, addin):
    """
    Loads or reloads the specifed module, which can be a string name
    or module object, then calls scan_module.

    Internal use only, users should prefer to import "xloil.importers"
    which hooks import/reload to trigger a module scan.
    """
    
    if isinstance(what, str):
        _module_addin_map[what] = addin
        module = importlib.import_module(what)
    elif inspect.ismodule(what):
        module = importlib.reload(what) # can we avoid calling our hooked reload?
    else:
        # We don't care about the return value currently
        result = []
        with StatusBar(3000) as status:
            for m in what:
                status.msg(f"Loading {m}")
                result.append(_import_scan(m))
        return result
    
    scan_module(module, addin)
    return module

def _import_file(path, addin, workbook_name=None):

    """
    Imports the specifed py file as a module without adding its path to sys.modules.

    Optionally also adds xlOil linked workbook name information.
    """

    with StatusBar(3000) as status:
        try:
            status.msg(f"Loading {path}...")
            directory, filename = os.path.split(path)
            filename = os.path.splitext(filename)[0]
            
            # avoid name collisions when loading workbook modules
            module_name = filename
            if workbook_name is not None:
                module_name = "xloil_wb_" + filename
                _linked_workbooks[module_name] = workbook_name

            if len(directory) > 0 or workbook_name is not None:
                _module_finder.path_map[module_name] = path

            _module_addin_map[module_name] = addin
            module = importlib.import_module(module_name)

            # Calling import_module will bypass our import hook, so scan_module explicitly
            scan_module(module)

            status.msg(f"Finished loading {path}")

            return module

        except Exception as e:

            log_except(f"Failed to load module {path}")
            status.msg(f"Error loading {path}, see log")


_real_builtin_import   = builtins.__import__
_real_importlib_reload = importlib.reload



def _import_hook(name, *args, **kwargs):

    #TODO: check if already in sys.modules

    module = _real_builtin_import(name, *args, **kwargs)

    # If name is of the form "foo.bar", the result of _real_builtin_import will point to 
    # the top level package, not the module we want to scan
    real_module = sys.modules.get(name, module)
    scan_module(real_module)

    return module

def _reload_hook(*args, **kwargs):
    _clear_pending_registrations(args[0])
    module = _real_importlib_reload(*args, **kwargs)
    scan_module(module)
    return module

builtins.__import__ = _import_hook
importlib.reload = _reload_hook

def _unhook_import():
    builtins.__import__ = _real_builtin_import
    importlib.reload    = _real_importlib_reload
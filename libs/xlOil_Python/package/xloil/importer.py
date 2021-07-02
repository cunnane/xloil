"""
Importing this module also hooks the import and reload functions
"""

import importlib
import importlib.util
import importlib.abc
import builtins
import sys
import os

from .shadow_core import *
from .xloil import scan_module, _clear_pending_registrations

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


# We maintain a _ModuleFinder on sys.meta_path to catch any reloads of our non-standard 
# loaded modules
_module_finder = _ModuleFinder()
sys.meta_path.append(_module_finder)

def import_from_file(path, workbook_name=None):

    """
    Imports the specifed py file as a module without adding its path to sys.modules.

    Optionally also adds xlOil linked workbook name information.
    """

    directory, filename = os.path.split(path)
    filename = filename.replace('.py', '')

    # avoid name collisions when loading workbook modules
    module_name = filename if workbook_name is None else "xloil_wb_" + filename

    if len(directory) > 0 or workbook_name is not None:
        _module_finder.path_map[module_name] = path
   
    module = importlib.import_module(module_name)

    # Allows 'local' modules to know which workbook they link to
    if workbook_name is not None:
        module._xloil_workbook = workbook_name
        module._xloil_workbook_path = os.path.join(directory, workbook_name)

    # Calling import_module will bypass our import hook, so scan_module explicitly
    scan_module(module)

    return module


#
# Hook 'import' and importlib.reload
#

_real_builtin_import = builtins.__import__

def _import_hook(name, *args, **kwargs):

    module = _real_builtin_import(name, *args, **kwargs)

    # If name is of the form "foo.bar", module will point to the top level 
    # package, not the module we want to scan
    real_module = sys.modules.get(name, module)
    scan_module(real_module)

    return module

builtins.__import__ = _import_hook


_real_importlib_reload =  importlib.reload

def _reload_hook(*args, **kwargs):
    _clear_pending_registrations(args[0])
    module = _real_importlib_reload(*args, **kwargs)
    scan_module(module)
    return module

importlib.reload = _reload_hook

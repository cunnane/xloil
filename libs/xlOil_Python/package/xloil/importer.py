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
from ._common import log, log_except

_module_addin_map = dict() # Keeps track of which addin loads a particular source file
_linked_workbooks = dict()

class _SpecifiedPathFinder(importlib.abc.MetaPathFinder):
    """
    Allows importing a module from a path specified in path_map without
    needing to add it to sys.paths - essentially a private set of import 
    paths, indexed by module name
    """

    _path_map = dict()

    def find_spec(self, fullname, path, target=None):
        path = self._path_map.get(fullname, None)
        if path is None:
            return None
        return importlib.util.spec_from_file_location(fullname, path)

    def find_module(self, fullname, path):
        return None

    def add_path(self, name, path):
        self._path_map[name] = path

# Install a sys.meta_path hook. This allows reloads to work for modules 
# we import from specific path in _import_file
_module_finder = _SpecifiedPathFinder()
sys.meta_path.append(_module_finder)


def linked_workbook(mod=None):
    """
        Returns the full path of the workbook linked to the specified module
        or None if the module was not loaded with an associated workbook.
        If no module is specified, the calling module is used.
    """
    if mod is None:
        # Get immediate caller
        frame = inspect.stack()[1]
    return _linked_workbooks.get(frame.filename, None)


def source_addin(mod=None):
    if mod is None:
        # Get the highest level caller we recognise
        for frame in inspect.stack()[::-1]:
            addin = _module_addin_map.get(frame.filename, None)
            if addin is not None:
                return addin
    return None


def _import_scan(what, addin):
    """
    Loads or reloads the specifed module, which can be a string name
    or module object, then calls scan_module.

    Internal use only.
    """
    
    if isinstance(what, str):
        module = importlib.import_module(what)
        # Remember which addin loaded this module
        _module_addin_map[module.__file__] = addin ## TODO: hasattr
    elif inspect.ismodule(what):
        module = importlib.reload(what)
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

    Optionally also adds xlOil linked workbook information.
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
                _linked_workbooks[path] = workbook_name

            if len(directory) > 0 or workbook_name is not None:
                _module_finder.add_path(module_name, path)

            _module_addin_map[path] = addin
            module = importlib.import_module(module_name)

            # Calling import_module will bypass our import hook, so scan_module explicitly
            scan_module(module)

            status.msg(f"Finished loading {path}")

            return module

        except Exception as e:

            log_except(f"Failed to load module {path}")
            status.msg(f"Error loading {path}, see log")


def _install_hook():
    # Hooks the import mechanism to run register.scan_module on all .py files.
    # We copy _bootstrap_external._install, replacing the source loader with one which 
    # runs scan_module and install our finder at the start of sys.path_hooks

    from importlib.machinery import (
        SOURCE_SUFFIXES, BYTECODE_SUFFIXES, FileFinder, 
        ExtensionFileLoader, SourcelessFileLoader, SourceFileLoader
        )
    import _imp

    class LoadAndScan(SourceFileLoader):
        def exec_module(self, module):
            super().exec_module(module)
            scan_module(module)

    extensions = ExtensionFileLoader, _imp.extension_suffixes()
    source     = LoadAndScan, SOURCE_SUFFIXES
    bytecode   = SourcelessFileLoader, BYTECODE_SUFFIXES

    sys.path_hooks.insert(0, FileFinder.path_hook(*[extensions, source, bytecode]))

_install_hook()
import importlib
import importlib.util
import importlib.abc
import builtins
import sys
import os
import inspect

from .register import scan_module, _clear_pending_registrations
from ._core import StatusBar, Addin
from .logging import log, log_except

_module_addin_map = dict() # Stores which addin loads a particular source file
_linked_workbooks = dict() # Stores the workbooks associated with an source file 

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

def _pump_message_loop(loop, timeout:float):
    """
    Called internally to run the asyncio message loop. Returns the number of active tasks
    """
    import asyncio

    async def wait():
        await asyncio.sleep(timeout)
    
    loop.run_until_complete(wait())

    all_tasks = asyncio.all_tasks if sys.version_info[:2] > (3, 6) else asyncio.Task.all_tasks
    return len([task for task in all_tasks(loop) if not task.done()])

def linked_workbook() -> str:
    """
        Returns the full path of the workbook linked to the calling module
        or None if the module was not loaded with an associated workbook.
    """
    # Get immediate caller
    frame = inspect.stack()[1]
    return _linked_workbooks.get(frame.filename, None)


def source_addin() -> Addin:
    """
        Returns the full path of the source add-in (XLL file) associated with
        the current code. That is the add-in which has caused the current code
        to be executed
    """
    import xloil_core

    addin_path = None

    # Get the highest level caller we recognise
    for frame in inspect.stack()[::-1]:
        addin_path = _module_addin_map.get(frame.filename, None)
        if addin_path is not None:
            break

    return xloil_core.core_addin() if addin_path is None \
        else xloil_core.xloil_addins[addin_path] 


def get_event_loop():
    """
        Returns the background *asyncio* event loop used to load the current add-in. 
        Unless specified in the settings, all add-ins are loaded in the same thread  
        and event loop.
    """
    return source_addin().event_loop


def _import_and_scan(what, addin):
    """
    Loads or reloads the specifed module, which can be a string name
    or module object, then calls scan_module.

    Internal use only.
    """
    
    if isinstance(what, str):
        # Remember which addin loaded this module
        _module_addin_map[what] = addin.pathname
        module = importlib.import_module(what)
    elif inspect.ismodule(what):
        module = importlib.reload(what)
    else:
        # We don't care about the return value currently
        result = []
        with StatusBar(2000) as status:
            for m in what:
                status.msg(f"Loading {m}")
                result.append(_import_and_scan(m, addin))
            status.msg("xlOil load complete")
        return result
    
    scan_module(module, addin)
    return module

def _import_file_and_scan(path, addin=None, workbook_name:str=None):

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

            addin = addin or source_addin()

            _module_addin_map[path] = addin.pathname

            # Force a reload if an attempt is made to load a module again.
            # This can happen if a workbook is closed and reopened - it is
            # difficult to get python to delete the module. Without a reload
            # the 'pending funcs' won't be populated for the registration 
            # machinery.
            try:
                module = importlib.reload(sys.modules[module_name])
            except KeyError:
                module = importlib.import_module(module_name)

            # Calling import_module will bypass our import hook, so scan_module explicitly
            n_funcs = scan_module(module, addin)

            status.msg(f"Registered {n_funcs} funcs for {path}")

            return module

        except Exception as e:

            status.msg(f"Error loading {path}, see log")
            raise ImportError(f"{str(e)} whilst loading {path}", path=path, name=module_name) from e

from importlib.machinery import SourceFileLoader

class _LoadAndScanHook(SourceFileLoader):
    def exec_module(self, module):
        global _module_addin_map

        # See if _import_and_scan has written addin info for this module
        addin = _module_addin_map.get(module.__name__)
        if addin is not None:
            _module_addin_map[module.__file__] = addin

        # Exec module as normal
        super().exec_module(module)

        # Look for xlOil functions to register
        scan_module(module)

def _install_hook():
    # Hooks the import mechanism to run register.scan_module on all .py files.
    # We copy _bootstrap_external._install, replacing the source loader with one which 
    # runs scan_module and install our finder at the start of sys.path_hooks

    from importlib.machinery import (
        SOURCE_SUFFIXES, BYTECODE_SUFFIXES, FileFinder, 
        ExtensionFileLoader, SourcelessFileLoader
        )
    import _imp

    extensions = ExtensionFileLoader, _imp.extension_suffixes()
    source     = _LoadAndScanHook, SOURCE_SUFFIXES
    bytecode   = SourcelessFileLoader, BYTECODE_SUFFIXES

    sys.path_hooks.insert(0, FileFinder.path_hook(*[extensions, source, bytecode]))

    importlib.invalidate_caches()
    sys.path_importer_cache.clear()

    log.debug("Installed importlib hook to call scan_module")

_install_hook()
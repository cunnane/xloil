import importlib
import importlib.util
import importlib.abc
from logging import root
import sys
import os
import inspect
from importlib.machinery import SourceFileLoader

from .register import scan_module
from ._core import StatusBar, Addin, XLOIL_EMBEDDED
from .logging import log, log_except

_module_addin_map = dict() # Stores which addin loads a particular source file
_linked_workbooks = dict() # Stores the workbooks associated with an source file 


class _UrlLoader(importlib.abc.FileLoader, importlib.abc.SourceLoader):
    """
    Loads a python module from a URL, then runs `scan_module` on the result
    """
    def get_data(self, path):
        
        url = self.get_filename()

        log.debug("Loading module name '%s' from URL '%s'", self.name, url)

        # The core already loads the contents onedrive/sharepoint URLs
        # so we first try to fetch that cached copy, then otherwise fetch
        # the URL in the normal way with 'requests' 
        try:
            from xloil_core import _get_onedrive_source
            preloaded = _get_onedrive_source(url)
            return preloaded.encode('utf-8')

        except:
            import requests
            response = requests.get(url)
            return response.text.encode('utf-8')

    def exec_module(self, module):
        # Exec module as normal
        super().exec_module(module)

        # Look for xlOil functions to register
        scan_module(module)


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

        loader = None 

        if path.startswith("http"):
            loader = _UrlLoader(fullname, path)

        log.debug("Found spec for '%s' with location '%s'", fullname, path)
        return importlib.util.spec_from_file_location(
            fullname, path, 
            loader=loader, submodule_search_locations=[])

    def find_module(self, fullname, path):
        return None

    def add_path(self, name, path):
        log.debug("Associating module name '%s' with path '%s'", name, path)
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


def _import_file(path, addin=None, workbook_name:str=None):

    """
    Imports the specifed py file as a module without adding its path to sys.modules.

    Optionally also adds xlOil linked workbook information.
    """

    root_path, filename = os.path.split(path)
    filestem = os.path.splitext(filename)[0]
            
    module_name = filestem
    
    # avoid name collisions with any installed python modules when loading a
    # workbook modules (e.g. if your workbook was called sys.xlsx)
    
    if len(root_path) > 0:
        if workbook_name is not None:
            module_name = "xloil_wb_" + filestem # Uniquify accross wb?
        _module_finder.add_path(module_name, path)

    addin = addin or source_addin()

    _module_addin_map[path] = addin.pathname

    if workbook_name is not None:
        _linked_workbooks[path] = workbook_name

    log.info("Importing module %s from file '%s' at '%s' for addin '%s'. Linked workbook '%s'", 
             module_name, filename, root_path, addin, workbook_name)

    # Force a reload if an attempt is made to load a module again.
    # This can happen if a workbook is closed and reopened - it is
    # difficult to get python to delete the module. Without a reload
    # the 'pending funcs' won't be populated for the registration 
    # machinery.
    if module_name in sys.modules:
        module = importlib.reload(sys.modules[module_name])
    else:
        module = importlib.import_module(module_name)

    return module


def _import_and_scan(what, addin):
    """
    Loads or reloads the specifed module, which can be a string name
    or module object, then calls scan_module.

    Internal use only: called from xlOil Core
    """
    try:
        if isinstance(what, str):
            # Remember which addin loaded this module
            _module_addin_map[what] = addin.pathname
            module = importlib.import_module(what)
        elif inspect.ismodule(what):
            module = importlib.reload(what)
        else:
            return _import_and_scan_mutiple(what, addin)
    except (ImportError, ModuleNotFoundError) as e:
        import sys
        raise ImportError(f"{e.msg} with sys.path={sys.path}") from e
    
    scan_module(module, addin)
    return module

def _import_and_scan_mutiple(module_names, addin):
    result = []
    success = True
    with StatusBar(2000) as status:
        for m in module_names:
            status.msg(f"Loading {m}")
            log.debug("Loading python module '%s' for addin '%s'", m, addin)
            try:
                result.append(_import_and_scan(m, addin))
            except Exception as e:
                log_except(f"Failed to load '{m}'")
                status.msg(f"Failed to load '{m}'. See log")
                success = False
        if success:
            status.msg("xlOil python module load complete")
    return result

def _import_file_and_scan(path, addin=None, workbook_name:str=None):
    """
        Internal use only: called from xlOil Core
    """

    with StatusBar(3000) as status:
        try:
            status.msg(f"Loading {path}...")
            module = _import_file(path, addin, workbook_name)

            # Calling import_module will bypass our import hook, so scan_module explicitly
            n_funcs = scan_module(module, addin)
            status.msg(f"Registered {n_funcs} funcs for {path}")

        except Exception as e:
            status.msg(f"Error loading {path}, see log")
            raise ImportError(f"{str(e)} whilst loading {path}", path=path) from e


def import_functions(source:str, names=None, as_names=None, addin:Addin=None, workbook_name:str=None) -> None:
    """
        Loads functions from the specified source and registers them in Excel. The functions
        do not have to be decorated, but are imported as if they were decorated with ``xloil.func``.
        So if the functions have typing annotations, they are respected where possible.

        This function provides an analogue of ``from X import Y as Z`` but with Excel UDFS.

        Note: registering a large number of Excel UDFs will impair the function name-lookup performance 
        (which is by linear search through the name table).

        Parameters
        ----------

        source: str
            A module name or full path name to the target py file

        names: [Iterable[str] | dict | str]
            If not provided, the specified module is imported and any ``xloil.func`` decorated 
            functions are registered, i.e. call ``xloil.scan_module``. 
            
            If a str or an iterable of str, xlOil registers only the specified names regardless of whether 
            they are decorated. 
            
            If it is a ``dict``, it is interpreted as a map of source names to registered function
            names, i.e.``names = keys(), as_names=values()``. 

            If it is the string '*', xlOil will try to register all callables in the specified module,
            including async functions and class constructors.

        as_names: [Iterable[str]]
            If provided, specifies the Excel function names to register in the same order as `names`.
            Should have the same length as `names`.  If this is omitted, functions are registered under
            their python names.

        addin:
            Optional xlOil.Addin which the registered functions are associated with. If ommitted the
            currently executing addin is used, or the Core addin if this cannot be determined.

        workbook_name: [str]
            Optional workbook associated with the registered functions.

    """
    addin = addin or source_addin()
    
    module = source if inspect.ismodule(source) else \
        _import_file(source, addin, workbook_name)

    if names is None: 
        scan_module(module, addin)
        return

    if isinstance(names, str):
        if names == "*":
            from inspect import getmembers, isfunction, isclass, iscoroutinefunction, isasyncgenfunction
            source_names = [x[0] for x in getmembers(module, lambda x: 
                isfunction(x) or isclass(x) or iscoroutinefunction(x) or isasyncgenfunction(x))]
            target_names = source_names
        else:
            source_names = [names]
            target_names = source_names if as_names is None \
                else ([as_names] if isinstance(as_names, str) else as_names)
    else:
        try:
            source_names = names.keys()
            target_names = names.values()
        except AttributeError:
            source_names = names
            target_names = as_names or source_names

    from xloil.register import _register_functions, func

    def get_spec(obj, name):
        spec = getattr(obj, '_xloil_spec', None) or func(obj, register=False)._xloil_spec
        spec.name = name
        return spec

    to_register = [
        get_spec(getattr(module, source_name), target_name)
        for source_name, target_name in zip(source_names, target_names)
    ]
   
    _register_functions(to_register, module, addin, append=True)



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

if XLOIL_EMBEDDED:
    _install_hook()
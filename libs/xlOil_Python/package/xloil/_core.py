import importlib.util
from ._paths import XLOIL_BIN_DIR, add_dll_path
import os
import sys

# Tests if we have been loaded from the XLL plugin which will have
# already injected the xloil_core module
XLOIL_EMBEDDED = importlib.util.find_spec("xloil_core") is not None
XLOIL_READTHEDOCS = 'READTHEDOCS' in os.environ

def _fix_module_for_docs(namespace, target, replace):
    """
        When sphinx autodoc reads python objects, it uses their __module__
        attribute to determine their fully-qualified name.  When importing
        from a hidden private implementation, we'd like to rename this 
        __module__ so the import appeared to come from the top level package
    """
    for name in list(namespace):
        val = namespace[name]
        if getattr(val, '__module__', None) == target:
            val.__module__ = replace

if XLOIL_READTHEDOCS:

    from .stubs import xloil_core
    sys.modules['xloil_core'] = xloil_core

    #
    # If we are not called from an xlOil embedded interpreter, some symbols are 
    # missing so we define stubs for them. OK, it's just one
    #
    workbooks = xloil_core.Workbooks()
    """
        Collection of all open workbooks as Workbook objects.
    
        Examples
        --------

            workbooks['MyBook'].path
            workbooks.active.path

    """

elif not XLOIL_EMBEDDED:
    # We try to load xlOil_PythonXY.pyd where XY is the python version
    # if we succeed, we fake an entry in sys.modules so that future 
    # imports of 'xloil_core' will work as expected.
    import importlib
    
    sys.path.append(XLOIL_BIN_DIR)

    ver = sys.version_info
    pyd_name = f"xlOil_Python{ver.major}{ver.minor}"
    mod = None
    try:
        with add_dll_path(XLOIL_BIN_DIR):
            mod = importlib.import_module(pyd_name)
    except ModuleNotFoundError as e:
        raise ModuleNotFoundError(f"Failed to load {pyd_name} with " +
            f"sys.path={sys.path} and PATH={os.environ['PATH']}")
    
    sys.path.pop()
    sys.modules['xloil_core'] = mod


from xloil_core import *
from xloil_core import _LogWriter

if XLOIL_READTHEDOCS:
    _fix_module_for_docs(locals(), xloil_core.__name__, 'xloil')
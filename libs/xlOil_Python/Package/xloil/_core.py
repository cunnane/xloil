import importlib.util
from ._paths import XLOIL_BIN_DIR, add_dll_path, _get_environment_strings
import os
import sys

# Tests if we have been loaded from the XLL plugin which will have
# already injected the xloil_core module
XLOIL_EMBEDDED = importlib.util.find_spec("xloil_core") is not None
XLOIL_READTHEDOCS = 'READTHEDOCS' in os.environ

if XLOIL_EMBEDDED:
    """
    This looks like hocus pocus, but if we don't do it Qt (and possibly others)
    will fail to find environment variables we set prior to even loading the 
    python3.dll. I suspect this is something to do with having different environment
    blocks per version of the C runtime. See discussion https://bugs.python.org/issue16633
    This seems like the easist workaround for now.
    """
    env_vars = _get_environment_strings()
    for name, val in env_vars.items():
        if not name.startswith("="):
            os.environ[name] = val

    try: 
        from xloil_core import _LogWriter
        from conda.activate import CmdExeActivator
        if os.path.exists(os.path.join(sys.prefix, 'conda-meta', 'history')):
            activation = CmdExeActivator().build_activate(sys.prefix)
            for name, val in activation['export_vars'].items():
                os.environ[name] = str(val)
    except ImportError:
        pass

elif not XLOIL_READTHEDOCS:

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
    except (ImportError, ModuleNotFoundError) as e:
        raise type(e)(f"Failed to load {pyd_name} with " +
            f"sys.path={sys.path} and XLOIL_BIN_DIR={XLOIL_BIN_DIR} and PATH={os.environ['PATH']}")

    sys.path.pop()
    sys.modules['xloil_core'] = mod


try:
    from xloil_core import *
    # These classes back singletons, so we want their docstrings but we don't 
    # want to suggest they are part of the API, hence the leading underscore
    from xloil_core import _LogWriter, _AddinsDict, _DateFormatList 
    
except ImportError:

    # Fallback to stubs
    from .stubs.xloil_core import *
    from .stubs.xloil_core import _LogWriter, _AddinsDict, _DateFormatList 
    
    # Not completely sure this part is necessary in stubs mode
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
   

if XLOIL_READTHEDOCS:
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

    _fix_module_for_docs(locals(), xloil_core.__name__, 'xloil')

class _ActiveWorksheets:
    def __getitem__(self, name):
        return active_workbook().worksheets[name]

worksheets = _ActiveWorksheets()
"""
    Collection of Worksheets of the active Workbook
    
    Examples
    --------

        worksheets['Sheet1']['A1'].value = 'Hello'

"""

def create_gui(*args, **kwargs) -> ExcelGUI:
    # DEPRECATED. Rather create the xloil.ExcelGUI object directly.

    import warnings
    warnings.warn("create_gui is deprecated, create the ExcelGUI object directly", 
                  DeprecationWarning, stacklevel=2)
    if 'mapper' in kwargs:
        kwargs['funcmap'] = kwargs.pop('mapper')
    return ExcelGUI(*args, **kwargs)


class Singleton(type):
    _instances = {}
    def __call__(cls, *args, **kwargs):
        if cls not in cls._instances:
            cls._instances[cls] = super(Singleton, cls).__call__(*args, **kwargs)
        return cls._instances[cls]
    

class StatusBarExecutor:
    """
    Executor which displays messages in Excel's status bar whilst running.
    Errors are logged rather than raised.
    """
    def __init__(self, timeout):
        self._timeout = timeout
        
    def map(self, func, *args, message, job_name: str):
        """
        args should be a parameter list of iterables of arguments to pass to *func*,
        the same as python's built-in map. *message* should be a function which takes
        the same arguments as *func* and returns a message to be displayed.
        
        Errors are logged rather than raised.
        """
        success = True
        from .logging import log_except
        
        with StatusBar(self._timeout) as status:
             
            for arg_tuple in zip(*args):
                msg = message(*arg_tuple)   
                status.msg(msg)
                try:
                    yield func(*arg_tuple)
                except Exception as err:
                    log_except(f"Failed {msg}")
                    success = False
                    yield err
                    
            if success:
                status.msg(f"{job_name} succeeded")
            else:
                status.msg(f"{job_name} failed (see log)")


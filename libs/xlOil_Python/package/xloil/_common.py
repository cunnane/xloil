import importlib.util
import traceback

# Tests if we have been loaded from the XLL plugin which will have
# already injected the xloil_core module
XLOIL_EMBEDDED = importlib.util.find_spec("xloil_core") is not None

if not XLOIL_EMBEDDED:
    # We try to load xlOil_PythonXY.pyd where XY is the python version
    # if we succeed, we fake an entry in sys.modules so that future 
    # imports of 'xloil_core' will work as expected.
    import importlib
    import sys
    ver = sys.version_info
    dll_name = f"xlOil_Python{ver.major}{ver.minor}"
    sys.modules['xloil_core'] = importlib.import_module(dll_name)

from xloil_core import LogWriter

log = LogWriter()
"""
    Instance of `xloil.LogWriter` which writes a log message to xlOil's log.  The level 
    parameter can be a integer constant from the ``logging`` module or one of the strings
    *error*, *warn*, *info*, *debug* or *trace*.

    Only messages with a level higher than the xlOil log level which is (initially) set
    in the xlOil settings file will be output to the log file. Trace output can only
    be seen with a debug build of xlOil.
"""

def log_except(msg, level='error'):
    """
       Logs '{msg}: {stack trace}' with a default level of 'error'
    """
    log(f"{msg}: {traceback.format_exc()}", level='error')
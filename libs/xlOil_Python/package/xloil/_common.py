import importlib.util
import traceback

# Tests if we have been loaded from the XLL plugin which will have
# already injected the xloil_core module
XLOIL_EMBEDDED = importlib.util.find_spec("xloil_core") is not None

XLOIL_HAS_CORE = XLOIL_EMBEDDED

if not XLOIL_HAS_CORE:
    # We try to load xlOil_PythonXY.pyd where XY is the python version
    # if we succeed, we fake an entry in sys.modules so that future 
    # imports of 'xloil_core' will work as expected.
    import importlib
    import sys
    ver = sys.version_info
    dll_name = f"xlOil_Python{ver.major}{ver.minor}"
    try:
        sys.modules['xloil_core'] = importlib.import_module(dll_name)
        XLOIL_HAS_CORE = True
    except OSError:
        pass

# We can proceed without xloil_core: _core.py will set up placeholders
# for all core functionality - useful for type checking and help generation

if XLOIL_HAS_CORE:
    from xloil_core import (  # pylint: disable=import-error
        LogWriter
    )

else:
    class LogWriter:

        """
            Writes a log message to xlOil's log.  The level parameter can be a level constant 
            from the `logging` module or one of the strings *error*, *warn*, *info*, *debug* or *trace*.

            Only messages with a level higher than the xlOil log level which is initially set
            to the value in the xlOil settings will be output to the log file. Trace output
            can only be seen with a debug build of xlOil.
        """
        def __call__(self, msg, level=20):
            pass

        @property
        def level(self):
            """
            Returns or sets the current log level. The returned value will always be an 
            integer corresponding to levels in the `logging` module.  The level can be
            set to an integer or one of the strings *error*, *warn*, *info*, *debug* or *trace*.
            """
            pass

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
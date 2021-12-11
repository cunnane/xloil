import importlib.util
import traceback

XLOIL_HAS_CORE = importlib.util.find_spec("xloil_core") is not None


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

"""
    Instance of `xloil.LogWriter` which writes a log message to xlOil's log.  The level 
    parameter can be a integer constant from the ``logging`` module or one of the strings
    *error*, *warn*, *info*, *debug* or *trace*.

    Only messages with a level higher than the xlOil log level which is (initially) set
    in the xlOil settings file will be output to the log file. Trace output can only
    be seen with a debug build of xlOil.
"""
log = LogWriter()

def log_except(msg, level='error'):
    """
       Logs '{msg}: {stack trace}' with a default level of 'error'
    """
    log(f"{msg}: {traceback.format_exc()}", level='error')
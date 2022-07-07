import traceback
from . import _core

from xloil_core import _LogWriter

log = _LogWriter()
"""
    An instance of `xloil.LogWriter` which writes a log message to xlOil's log.  The level 
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
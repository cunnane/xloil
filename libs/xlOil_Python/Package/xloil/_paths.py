import sys
import os

ADDIN_NAME        = "xlOil.xll"
INIFILE_NAME      = "xlOil.ini"
APP_DATA_DIR      = os.path.join(os.environ.get('APPDATA','.'), "xlOil")
XLOIL_INSTALL_DIR = os.path.join(sys.prefix, "share", "xloil")
XLOIL_BIN_DIR     = os.environ.get("XLOIL_BIN_DIR", XLOIL_INSTALL_DIR)

class _SetPathContext:

    def __init__(self, path):
        self._path = path
        self._old_PATH = os.environ['PATH']

    def __enter__(self):
        os.environ['PATH'] += os.pathsep + self._path 
        return self

    def close(self):
        os.environ['PATH'] = self._old_PATH

    def __exit__(self, *args):
        self.close()

def add_dll_path(path):
    """
        Returns a context manager which adds a PATH to the dll search directory
        (either via AddDllDirectory or by changing PATH). The change is reversed
        on context exit
    """
    try:
        return os.add_dll_directory(path)
    except (AttributeError, FileNotFoundError):
        # Either Py < 3.8 or path does not exist (but we have to return a context mgr)
        return _SetPathContext(path)

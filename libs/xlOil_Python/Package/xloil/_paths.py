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

def _get_environment_strings():
    """
        Does the same as:
            import win32profile
            win32profile.GetEnvironmentStrings()
        But avoids the depedency on pywin32
    """
   
    import ctypes
    import locale

    kernel_func = ctypes.windll.kernel32.GetEnvironmentStringsA
    char_ptr = ctypes.POINTER(ctypes.c_char)
    kernel_func.restype = char_ptr

    # P will point to an block of char formatted as:
    #     name1=val1/0
    #     ...
    #     nameN=valN/0/0
    # (line-breaks added for clarity, they aren't present in the string)
    p = kernel_func()
    try:
    
        result = {}
        start = end = 0
        null = b'\x00' # Null-terminator for C-strings

        while True:
            while p[end] != null:
                end += 1
            if end == start:
                break
                
            keyval = p[start:end].decode(locale.getpreferredencoding()).split('=')
            # GetEnvironmentStrings returns some strange entries starting with '='
            if any(keyval[0]):
                result[keyval[0]] = keyval[1]
                
            end += 1    # Step over null terminator
            start = end # Move string start pointer
    finally:
        ctypes.windll.kernel32.FreeEnvironmentStringsA(p)
    
    return result
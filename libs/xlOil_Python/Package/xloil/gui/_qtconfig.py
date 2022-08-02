"""
    Do not import this module directly
"""

_QT_MODULE_NAME = ""

def QT_IMPORT(submod:str = None):
    """
        Allows switching between PyQt5 and PySide2 at runtime
    """
    global _QT_MODULE_NAME
    if len(_QT_MODULE_NAME) == 0:
        raise ImportError("Qt package not specifed, import pyqt5 or pyside2 first")

    import importlib
    return importlib.import_module(_QT_MODULE_NAME if submod is None else f"{_QT_MODULE_NAME}.{submod}")
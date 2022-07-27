"""
    Sets up xloil to use the PySide2 GUI toolkit. This *must* be imported before `PySide2`
    is imported to allow xlOil to own the QApplication object. *All* interaction with 
    the Qt GUI must be done on the GUI thread: use `Qt_thread().submit(...)`
"""

from . import _qtconfig

_qtconfig._QT_MODULE_NAME = "PySide2"

from ._qtgui import *

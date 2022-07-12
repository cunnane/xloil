"""
    Sets up xloil to use the PyQt5 GUI toolkit. This *must* be imported before `PyQt5`
    is imported to allow xlOil to own the QApplication object. *All* interaction with 
    the Qt GUI must be done on the GUI thread: use `Qt_thread().submit(...)`
"""

from . import _qtconfig

_qtconfig._QT_MODULE_NAME = "PyQt5"

from ._qtgui import *

# Trigger thread creation on import
Qt_thread()

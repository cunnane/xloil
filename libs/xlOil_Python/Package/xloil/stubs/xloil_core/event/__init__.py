"""
          Module containing event objects which can be hooked to receive events driven by 
          Excel's UI. The events correspond to COM/VBA events and are described in detail
          in the Excel Application API.
        
          Event Class
          -----------

              * Events are hooked using `+=`, e.g. `event.NewWorkbook += lambda wb: print(wb_name)`
              * Events are unhooked using `-=` passing a reference to the handler function
              * Each event has a `handlers` property listing all currently hooked handlers

          Events
          ------
          
              * AfterCalculate: 
                  Called after a calculation whether or not it completed or was interrupted
              * CalcCancelled:
                  Called when the user interrupts calculation by interacting with Excel.
              * WorkbookAfterClose:
                  Excel's event *WorkbookBeforeClose*, is  cancellable by the user so it is not 
                  possible to know if the workbook actually closed.  When xlOil calls 
                  `WorkbookAfterClose`, the workbook is certainly closed, but it may be some time
                  since that closure happened.
                  The event is not called for each workbook when xlOil exits.
              * PyBye:
                  An event fired just before xlOil finalises its embedded python interpreter. 
                  All python and xlOil functionality is still available. This event is useful 
                  to stop threads as it is called before threading module teardown, whereas 
                  `atexit` is called afterward.
              * UserException:
                  An event fired when an exception is raised in a user-supplied 
                  python callback, for example a GUI callback or and RTD publisher. 

          For other events see  `Excel.Application <https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)#events>`_
          
          Notes
          -----

              * The `CalcCancelled` and `WorkbookAfterClose` event are not part of the 
                Application object, see their individual documentation.
              * Where an event has reference parameter, for example the `cancel` bool in
                `WorkbookBeforeSave`, you need to set the value using `cancel.value=True`.
                This is because python does not support reference parameters for primitive types.


          Examples
          --------

          ::

              def greet(workbook, worksheet):
                  xlo.Range(f"[{workbook}]{worksheet}!A1") = "Hello!"

              xlo.event.WorkbookNewSheet += greet
              ...
              xlo.event.WorkbookNewSheet -= greet
              
              print(xlo.event.WorkbookNewSheet.handlers) # Should be empty

        """
from __future__ import annotations
import typing

__all__ = [
    "AfterCalculate",
    "ComAddinsUpdate",
    "Event",
    "NewWorkbook",
    "PyBye",
    "SheetActivate",
    "SheetBeforeDoubleClick",
    "SheetBeforeRightClick",
    "SheetCalculate",
    "SheetChange",
    "SheetDeactivate",
    "SheetSelectionChange",
    "UserException",
    "WorkbookActivate",
    "WorkbookAddinInstall",
    "WorkbookAddinUninstall",
    "WorkbookAfterClose",
    "WorkbookAfterSave",
    "WorkbookBeforeClose",
    "WorkbookBeforePrint",
    "WorkbookBeforeSave",
    "WorkbookDeactivate",
    "WorkbookNewSheet",
    "WorkbookOpen",
    "WorkbookRename",
    "XllAdd",
    "XllRemove",
    "allow",
    "boolRef",
    "pause"
]


class Event():
    def __iadd__(self, arg0: object) -> Event: ...
    def __isub__(self, arg0: object) -> Event: ...
    def clear(self) -> None: ...
    @property
    def handlers(self) -> tuple:
        """
        :type: tuple
        """
    pass
class boolRef():
    @property
    def value(self) -> bool:
        """
        :type: bool
        """
    @value.setter
    def value(self, arg1: bool) -> None:
        pass
    pass
def allow() -> None:
    """
    Resumes Excel's event handling after a pause.  Equivalent to VBA's
    `Application.EnableEvents = True` or `xlo.app().enable_events = True` 
    """
def pause() -> None:
    """
    Pauses Excel's event handling. Equivalent to VBA's 
    `Application.EnableEvents = False` or `xlo.app().enable_events = False` 
    """
AfterCalculate: xloil_core.event.Event = None
ComAddinsUpdate: xloil_core.event.Event = None
NewWorkbook: xloil_core.event.Event = None
PyBye: xloil_core.event.Event = None
SheetActivate: xloil_core.event.Event = None
SheetBeforeDoubleClick: xloil_core.event.Event = None
SheetBeforeRightClick: xloil_core.event.Event = None
SheetCalculate: xloil_core.event.Event = None
SheetChange: xloil_core.event.Event = None
SheetDeactivate: xloil_core.event.Event = None
SheetSelectionChange: xloil_core.event.Event = None
UserException: xloil_core.event.Event = None
WorkbookActivate: xloil_core.event.Event = None
WorkbookAddinInstall: xloil_core.event.Event = None
WorkbookAddinUninstall: xloil_core.event.Event = None
WorkbookAfterClose: xloil_core.event.Event = None
WorkbookAfterSave: xloil_core.event.Event = None
WorkbookBeforeClose: xloil_core.event.Event = None
WorkbookBeforePrint: xloil_core.event.Event = None
WorkbookBeforeSave: xloil_core.event.Event = None
WorkbookDeactivate: xloil_core.event.Event = None
WorkbookNewSheet: xloil_core.event.Event = None
WorkbookOpen: xloil_core.event.Event = None
WorkbookRename: xloil_core.event.Event = None
XllAdd: xloil_core.event.Event = None
XllRemove: xloil_core.event.Event = None

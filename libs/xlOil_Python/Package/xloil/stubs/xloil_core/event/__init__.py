"""
          A module containing event objects which can be hooked to receive events driven by 
          Excel's UI. The events correspond to COM/VBA events and are described in detail
          in the Excel Application API.
        
          See :ref:`Events:Introduction` and 
          `Excel.Application <https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)#events>`_

          Using the Event Class
          ---------------------

              * Events are hooked using `+=`, e.g.

              ::
              
                  event.NewWorkbook += lambda wb: print(wb_name)

              * Events are unhooked using `-=` and passing a reference to the handler function

              ::

                  event.NewWorkbook += foo
                  event.NewWorkbook -= foo

              * You should not return anything from an event handler

              * Each event has a `handlers` property listing all currently hooked handlers

              * Where an event has reference parameter, for example the `cancel` bool in
                `WorkbookBeforePrint`, you need to set the value using `cancel.value=True`.
                This is because python does not support reference parameters for primitive types.

                ::

                    def no_printing(wbName, cancel):
                      cancel.value = True
                    xlo.event.WorkbookBeforePrint += no_printing

              * Workbook and worksheet names are passed a string, Ranges as passed as a 
                :any:`xloil.Range`
    
          Python-only Events
          ------------------

          These events specific to python and not documented in the Core documentation:

            * PyBye:
                Fired just before xlOil finalises its embedded python interpreter. 
                All python and xlOil functionality is still available. This event is useful 
                to stop threads as it is called before threading module teardown, whereas 
                python's `atexit` is called afterwards. Has no parameters.
            * UserException:
                Fired when an exception is raised in a user-supplied python callback, 
                for example a GUI callback or an RTD publisher. Has no parameters.

          Examples
          --------

          ::

              def greet(workbook, worksheet):
                  xlo.Range(f"[{workbook}]{worksheet}!A1") = "Hello!"

              xlo.event.WorkbookNewSheet += greet
              ...
              xlo.event.WorkbookNewSheet -= greet
              
              print(xlo.event.WorkbookNewSheet.handlers) # Should be empty


          ::

              def click_handler(sheet_name, target, cancel):
                  xlo.worksheets[sheet_name]['A5'].value = target.address()
    
              xlo.event.SheetBeforeDoubleClick += click_handler

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
AfterCalculate: Event=None # value = <xloil_core.event.Event object>
ComAddinsUpdate: Event=None # value = <xloil_core.event.Event object>
NewWorkbook: Event=None # value = <xloil_core.event.Event object>
PyBye: Event=None # value = <xloil_core.event.Event object>
SheetActivate: Event=None # value = <xloil_core.event.Event object>
SheetBeforeDoubleClick: Event=None # value = <xloil_core.event.Event object>
SheetBeforeRightClick: Event=None # value = <xloil_core.event.Event object>
SheetCalculate: Event=None # value = <xloil_core.event.Event object>
SheetChange: Event=None # value = <xloil_core.event.Event object>
SheetDeactivate: Event=None # value = <xloil_core.event.Event object>
SheetSelectionChange: Event=None # value = <xloil_core.event.Event object>
UserException: Event=None # value = <xloil_core.event.Event object>
WorkbookActivate: Event=None # value = <xloil_core.event.Event object>
WorkbookAddinInstall: Event=None # value = <xloil_core.event.Event object>
WorkbookAddinUninstall: Event=None # value = <xloil_core.event.Event object>
WorkbookAfterClose: Event=None # value = <xloil_core.event.Event object>
WorkbookAfterSave: Event=None # value = <xloil_core.event.Event object>
WorkbookBeforeClose: Event=None # value = <xloil_core.event.Event object>
WorkbookBeforePrint: Event=None # value = <xloil_core.event.Event object>
WorkbookBeforeSave: Event=None # value = <xloil_core.event.Event object>
WorkbookDeactivate: Event=None # value = <xloil_core.event.Event object>
WorkbookNewSheet: Event=None # value = <xloil_core.event.Event object>
WorkbookOpen: Event=None # value = <xloil_core.event.Event object>
WorkbookRename: Event=None # value = <xloil_core.event.Event object>
XllAdd: Event=None # value = <xloil_core.event.Event object>
XllRemove: Event=None # value = <xloil_core.event.Event object>

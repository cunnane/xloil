"""
          A module containing event objects which can be hooked to receive events driven by 
          Excel's UI. The events correspond to COM/VBA events and are described in detail
          in the Excel Application API. The naming convention (including case) of the VBA events
          has been preserved for ease of search.
          
        
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

          These events are specific to python and not noted in the Core documentation:

            * PyBye:
                Fired just before xlOil finalises its embedded python interpreter. 
                All python and xlOil functionality is still available. This event is useful 
                to stop threads as it is called before threading module teardown, whereas 
                python's `atexit` is called afterwards. Has no parameters.
            * UserException:
                Fired when an exception is raised in a user-supplied python callback, 
                for example a GUI callback or an RTD publisher. Has no parameters.
            * file_change:
                This is a special parameterised event, see the separate documentation
                for this function.

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
    "file_change",
    "pause"
]


class Event():
    def __iadd__(self, arg0: object) -> Event: ...
    def __isub__(self, arg0: object) -> Event: ...
    def add(self, handler: object) -> Event: 
        """
        Registers an event handler callback with this event, equivalent to
        `event += handler`
        """
    def clear(self) -> None: 
        """
        Removes all handlers from this event
        """
    def remove(self, handler: object) -> Event: 
        """
        Deregisters an event handler callback with this event, equivalent to
        `event -= handler`
        """
    @property
    def handlers(self) -> tuple:
        """
        The tuple of handlers registered for this event. Read-only.

        :type: tuple
        """
    pass
class _bool_ref():
    @property
    def value(self) -> bool:
        """
        :type: bool
        """
    @value.setter
    def value(self, arg1: bool) -> None:
        pass
    pass
def allow(excel: bool = True) -> None:
    """
    Resumes event handling after a previous call to *pause*.

    If *excel* is True (the default), also calls `Application.EnableEvents = True`
    (equivalent to `xlo.app().enable_events = True`)
    """
def file_change(path: str, action: str = 'modify', subdirs: bool = True) -> Event:
    """
    This function returns an event specific to the given path and action; the 
    event fires when a watched file or directory changes.  The returned event 
    can be hooked in the usual way using `+=` or `add`. Calling this function 
    with same arguments always returns a reference to the same event object.

    The handler should take a single string argument: the name of the file or 
    directory which changed.

    The event runs on a background thread.

    Parameters
    ----------

    path: str
       Can be a file or a directory. If *path* points to a directory, any change
       to files in that directory, will trigger the event. Changes to the specified 
       directory itself will not trigger the event.

    action: str ["add", "remove", "modify"], default "modify"
       The event will only fire when this type of change is detected:

         *modify*: any update which causes a file's last modified time to change
         *remove*: file deletion
         *add*: file creation

       A file rename triggers *remove* followed by *add*.

    subdirs: bool (true)
       including in subdirectories,
    """
def pause(excel: bool = True) -> None:
    """
    Stops all xlOil event handling - any executing handlers will complete but
    no further handlers will fire.

    If *excel* is True (the default), also calls `Application.EnableEvents = False`
    (equivalent to `xlo.app().enable_events = False`)
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

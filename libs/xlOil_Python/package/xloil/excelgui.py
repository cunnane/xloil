from ._core import TaskPaneFrame, ExcelGUI, ExcelWindow
from . import _core


_task_panes = set()
 
class CustomTaskPane:
    """
        Base class for custom task panes. Can be sub-classed to implement
        task panes with different GUI toolkits.  Subclasses should implement
        at least the `on_visible` and `on_size` events.
    """

    def __init__(self, pane: TaskPaneFrame):
        self._pane = pane
        self._pane.add_event_handler(self)
        _task_panes.add(self)

    def on_size(self, width, height):
        """
        Called when the task pane is resized
        """
        ...
             
    def on_visible(self, value):
        """ Called when the visible state changes, passes the new state """
        ...

    def on_docked(self):
        """ Called when the pane is docked to a new location or undocked """
        ...

    def on_destroy(self):
        """ Called before the pane is destroyed to release any resources """

        # Release internal task pane pointer
        self._pane = None
        # Remove ourselves from pane lookup table
        _task_panes.remove(self)

    @property
    def pane(self) -> TaskPaneFrame:
        """Returns the TaskPaneFrame a reference to the window holding the python GUI"""
        return self._pane

    @property
    def visible(self) -> bool:
        """Returns True if the pane is currently shown"""
        return self._pane.visible

    @visible.setter
    def visible(self, value: bool):
        self._pane.visible = value

    @property
    def size(self) -> tuple:
        """Returns a tuple (width, height)"""
        return self._pane.size

    @size.setter
    def size(self, value:tuple):
        self._pane.size = value


def find_task_pane(title:str=None, workbook=None, window=None):
    """
        Finds all xlOil python task panes associated with the active window, 
        optionally filtering by the pane `title`. 

        This primary use of this function is to look for an existing task pane
        before creating a new one.

        Task panes are linked to Excel's Window objects which can have a many-to-one
        relationship with workbooks. If a `workbook` name is specified, all task panes 
        associated with that workbook will be searched.

        Returns: if `title` is specified, returns a (case-sensitive) match of a single
        `xloil.CustomTaskPane object` or None, otherwise returns a list of 
        `xloil.CustomTaskPane` objects.
    """
    from . import _core

    if window is None:
        if workbook is None:
            hwnds = [_core.windows.active.hwnd]
        else:
            workbook = _core.Workbook(workbook) # TODO: string or obj....
            hwnds = [x.hwnd for x in _core.windows]
    else:
            hwnds = [ExcelWindow(window).hwnd]

    found = [x for x in _task_panes if x.pane.window.hwnd in hwnds]
    if title is None:
        return found
    else:
        return next((x for x in found if x.pane.title == title), None)


async def create_task_pane(
    name:str, creator=None, window=None, gui:ExcelGUI=None, size=None, visible=True):
    """
        Returns a task pane with title <name> attached to the active window,
        creating it if it does not already exist.  This function is equivalent
        to `ExcelGUI.create_task_pane(...)`

        Parameters
        ----------

        creator: 
            * a subclass of `QWidget` or
            * a function which takes a `TaskPaneFrame` and returns a `CustomTaskPane`

        window: 
            a window title or `ExcelWindow` object to which the task pane should be
            attached.  If None, the active window is used.

        gui: `ExcelGUI` object
            GUI context used when creating a pane.    

    """
    pane = find_task_pane(name)
    if pane is not None:
        pane.visible = visible
        return pane
    if creator is None:
        return None

    from .qtgui import _try_create_qt_pane
    creator = _try_create_qt_pane(creator) or creator

    if isinstance(window, str):
        window = ExcelWindow(window)

    frame = await gui.task_pane_frame(name, window)
    pane = creator(frame)

    pane.visible = visible
    if size is not None:
        pane.size = size

    return pane
from .shadow_core import TaskPaneFrame, windows

_task_panes = set()

class CustomTaskPane:
    """
        Base class for custom task pane event handler. Can be sub-classes to 
        implement task panes with different GUI toolkits
    """

    def __init__(self, pane: TaskPaneFrame):
        self._pane = pane
        self._pane.add_event_handler(self)
        _task_panes.add(self)

    def on_size(self, width, height):
        # Called when the task pane is resized
        ...
             
    def on_visible(self, value):
        # Called when the visible state changes, passes the new state
        ...

    def on_docked(self):
        # Called when the pane is docked to a new location or undocked
        ...

    def on_destroy(self):
        # Called before the pane is destroyed to release any resources

        # Release internal task pane pointer
        self._pane = None
        # Remove ourselves from pane lookup table
        _task_panes.remove(self)

    @property
    def pane(self):
        return self._pane

    @property
    def visible(self):
        return self._pane.visible

    @visible.setter
    def visible(self, value: bool):
        self._pane.visible = value


def find_task_pane(title=None, workbook=None):
    """
        Finds any xlOil python task panes associated with the active window, 
        optionally filtering by pane `title`. 

        This function should be used to find an existing task pane to make
        visible before a task pane is created.

        Task panes are linked to Excel's Window objects which can have a many-to-one
        relationship with workbooks. If a `workbook` name is specified, all task panes 
        associated with that workbook will be searched.

        Returns: if `title` is specified, returns a (case-sensitive) match of a single
        `xloil.CustomTaskPane object` or None, otherwise returns a list of 
        `xloil.CustomTaskPane` objects.
    """

    hwnds = [x.hwnd for x in windows]
    found = [x for x in _task_panes if x.pane.window.hwnd in hwnds]
    if title is None:
        return found
    else:
        return next((x for x in found if x.pane.title == title), None)

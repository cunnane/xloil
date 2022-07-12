from xloil import TaskPaneFrame, ExcelGUI, ExcelWindow, Workbook
import concurrent.futures as futures
import concurrent.futures.thread
import xloil

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

    if window is None:
        if workbook is None:
            hwnds = [xloil.app().windows.active.hwnd]
        else:
            workbook = Workbook(workbook) # TODO: string or obj....
            hwnds = [x.hwnd for x in xloil.app().windows]
    else:
            hwnds = [ExcelWindow(window).hwnd]

    found = [x for x in _task_panes if x.pane.window.hwnd in hwnds]
    if title is None:
        return found
    else:
        return next((x for x in found if x.pane.title == title), None)


def _try_create_qt_pane(obj):

    from ._qtgui import QtThreadTaskPane, QT_IMPORT
    QWidget = QT_IMPORT("QtWidgets").QWidget
    if issubclass(obj, QWidget):
        return lambda pane: QtThreadTaskPane(pane, obj)
    return None


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

    creator = _try_create_qt_pane(creator) or creator

    if isinstance(window, str):
        window = ExcelWindow(window)

    frame = await gui.task_pane_frame(name, window)
    pane = creator(frame)

    pane.visible = visible
    if size is not None:
        pane.size = size

    return pane


class _GuiExecutor(futures.Executor):
    """
        GUIs toolkits like to be accessed from the same thread consistently.
        This base class creates a thread and manages a command queue.  
        Subclasses must define a `_main` method to initialise the toolkit 
        and a `_shutdown` method to destroy it.
    """

    def __init__(self, name):
        import threading
        import queue

        self._work_queue = queue.SimpleQueue()
        self._thread = threading.Thread(target=self._main_loop, name=name)
        self._broken = False
        self._thread.start()

        # PyBye is called before `threading` module teardown, whereas `atexit` comes later
        xloil.event.PyBye += self.shutdown

    def _do_work(self):
        import queue
        try:
            while True:
                work_item = self._work_queue.get_nowait()
                if work_item is not None:
                    work_item.run()
                    del work_item
        except queue.Empty:
            return
        except Exception as e:
             xloil.log(f"{self._thread.name} error running job: {e}", level='warn')


    def submit(self, fn, *args, **kwargs):
        if self._broken:
            raise futures.BrokenExecutor(self._broken)

        f = futures.Future()
        w = concurrent.futures.thread._WorkItem(f, fn, args, kwargs)

        self._work_queue.put(w)

        return f

    def shutdown(self, wait=True, cancel_futures=False):
        # TODO: implement wait/cancel_futures support?
        if not self._broken:
            self.submit(self._shutdown)

    def _main_loop(self):

        try:
            self._main()
        except Exception as e:
            xloil.log(f"{self._thread.name} failed: {e}", level='error')

        self._broken = True

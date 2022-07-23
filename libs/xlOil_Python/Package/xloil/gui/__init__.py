from xloil import TaskPaneFrame, ExcelGUI, ExcelWindow, Workbook
import concurrent.futures as futures
import concurrent.futures.thread
import xloil
import asyncio
import typing

_task_panes = set()

async def create_gui(ribbon=None, func_names=None, name=None) -> ExcelGUI:
    """
        Returns a coroutine which returns an ExcelGUI object. The *ExcelGUI*
        is connected using the specified ribbon customisation XML to Excel.  
        When the *ExcelGUI* object is deleted, it unloads the associated COM 
        add-in and so all Ribbon customisation and attached task panes.

        Parameters
        ----------

        ribbon: str
            A Ribbon XML string, most easily created with a specialised editor.
            The XML format is documented on Microsoft's website

        func_names: Func[str -> callable] or Dict[str, callabe]
            The ``func_names`` mapper links callbacks named in the Ribbon XML to
            python functions. It can be either a dictionary containing named 
            functions or any callable which returns a function given a string.
            Each return handler should take a single ``RibbonControl``
            argument which describes the control which raised the callback.

            Callbacks declared async will be executed in the addin's event loop. 
            Other callbacks are executed in Excel's main thread. Async callbacks 
            cannot return values.

        name: str
            The addin name which will appear in Excel's COM addin list.
            If None, uses the filename at the call site as the addin name.

    """

    if name is None:
        import inspect
        import os
        caller_frame = inspect.stack()[1]
        filepath = caller_frame.filename

        # get rid of the directory and split filename and extension
        name = os.path.splitext(os.path.basename(filepath))[0]

    # Does nothing until connected
    gui = ExcelGUI(name)

    async def connect():
        await gui.connect(ribbon, func_names)
        return gui

    return await connect()


class CustomTaskPane:
    """
        Base class for custom task panes. xlOil provides two toolkit-specfic
        implementations: `xloil.gui.pyqt5.QtThreadTaskPane` (pyside is also 
        supported) and `xloil.gui.tkinter.TkThreadTaskPane`.
        
        Can be sub-classed to implement task panes with different GUI toolkits.
    """

    def __init__(self):
        pass

    def _attach_frame(self, frame: TaskPaneFrame):
        """
            Attaches this *CustomTaskPane* to a *TaskPaneFrame*  causing it 
            to be resized/moved with the pane window.  Called by 
            `xloil.ExcelGUI.attach_frame` and generally should not need to be 
            called directly.
        """
        self._frame = frame
        _task_panes.add(self)
        frame.attach(self, self.hwnd())

    def draw(self):
        raise NotImplementedError()

    def on_visible(self, value):
        """ Called when the visible state changes, passes the new state """
        # TODO: no longer necessary
        ...

    def on_docked(self):
        """ Called when the pane is docked to a new location or undocked """
        ...

    def on_destroy(self):
        """ Called before the pane is destroyed to release any resources """

        # Release internal task pane pointer
        self._frame = None
        # Remove ourselves from pane lookup table
        _task_panes.remove(self)

    @property
    def frame(self) -> TaskPaneFrame:
        """Returns the TaskPaneFrame: a reference to the window holding the python GUI"""
        return self._frame

    @property
    def visible(self) -> bool:
        """Returns True if the pane is currently shown"""
        return self._frame.visible

    @visible.setter
    def visible(self, value: bool):
        self._frame.visible = value

    @property
    def size(self) -> tuple:
        """Returns a tuple (width, height)"""
        return self._frame.size

    @size.setter
    def size(self, value:tuple):
        self._frame.size = value


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
            if isinstance(workbook, str):
                workbook = Workbook(workbook)
            hwnds = [x.hwnd for x in workbook.windows]
    else:
        if isinstance(window, str):
            window = ExcelWindow(window)
        hwnds = [window.hwnd]

    found = [x for x in _task_panes if x.frame.window.hwnd in hwnds]
    if title is None:
        return found
    else:
        return next((x for x in found if x.frame.title == title), None)


async def find_or_create_pane(
        name: str, 
        creator: typing.Callable[[], typing.Awaitable[CustomTaskPane]],
        size=None, 
        visible=True) -> CustomTaskPane:
    """
        Returns a task pane with title <name> attached to the active window,
        creating it if it does not already exist.  This function is equivalent
        to `ExcelGUI.create_task_pane(...)`

        Parameters
        ----------

        name: 
            The name of the pane to find or create.

        creator: 
            An async function which takes no arguments and returns (a coroutine to)
            a *CustomTaskPane*. This function is called if the named pane is not
            found and typically is something like
            
            ::

                lambda: gui.attach_pane(name, QtThreadTaskPane(OurQWidget), window)
            
        size:
            If provided, a tuple (width, height) used to set the pane size

        visible:
            Determines the pane visibility. Defaults to True.

    """
    pane = find_task_pane(name)
    if pane is not None:
        pane.visible = visible
        return pane

    pane = await creator()

    pane.visible = visible
    if size is not None:
        pane.size = size

    return pane


def _try_create_from_qwidget(obj):
    """
        If obj is a QWidget, returns a creator function which uses QtThreadTaskPane.
    """
    try:
        # This will raise an ImportError if Qt has not been properly
        # setup by importing xloil.gui.pyqt5 (or pyside2). This check is
        # very cheap: _qtconfig imports nothing.
        from ._qtconfig import QT_IMPORT
        QWidget = QT_IMPORT("QtWidgets").QWidget
        
        if isinstance(obj, type) and issubclass(obj, QWidget) or isinstance(obj, QWidget):
            from ._qtgui import QtThreadTaskPane
            return QtThreadTaskPane(obj)

    except ImportError:
        pass

    return obj


async def _attach_task_pane(
        gui: ExcelGUI,
        name: str, 
        pane: CustomTaskPane,
        window: ExcelWindow, 
        size: tuple, 
        visible: bool):
    
    # Implementation of ExcelGUI.attach_pane

    if isinstance(window, str):
        window = ExcelWindow(window)
    
    pane = _try_create_from_qwidget(pane)

    frame = await gui._create_task_pane_frame(name, window)
    pane._attach_frame(frame)

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
        if not self._broken: # TODO: implement wait/cancel_futures support?
            self.submit(self._shutdown)

    def _main_loop(self):

        try:
            self._main()
        except Exception as e:
            xloil.log(f"{self._thread.name} failed: {e}", level='error')

        self._broken = True

    def _wrap(self, fn, discard=False):
        """
            Called by Tk_thread, Qt_thread to implement their decorator 
            behaviour.
        """
        if discard:
            def logged(*arg, **kwargs):
                try:
                    return fn(*args, **kwargs)
                except Exception as e:
                    xlo.log("Error running job on GUI thread: {str(e)}", level="error")
            fn = logged

        def wrapped(*args, **kwargs):
            return self.submit(fn, *args, **kwargs)

        return wrapped
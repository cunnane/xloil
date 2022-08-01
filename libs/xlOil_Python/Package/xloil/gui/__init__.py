from xloil import TaskPaneFrame, ExcelGUI, ExcelWindow, Workbook

import concurrent.futures as futures
import concurrent.futures.thread
import xloil
import asyncio
import typing
import threading
import inspect

_task_panes = set()

class CustomTaskPane:
    """
        Base class for custom task panes. xlOil provides two toolkit-specfic
        implementations: `xloil.gui.pyqt5.QtThreadTaskPane` (pyside is also 
        supported) and `xloil.gui.tkinter.TkThreadTaskPane`.
        
        Can be sub-classed to implement task panes with different GUI toolkits.
    """

    def __init__(self):
        self.hwnd = None

    async def _attach_frame(self, frame: typing.Awaitable[TaskPaneFrame]):
        """
            Attaches this *CustomTaskPane* to a *TaskPaneFrame*  causing it 
            to be resized/moved with the pane window.  Called by 
            `xloil.ExcelGUI.attach_frame` and generally should not need to be 
            called directly.
        """
        self.hwnd = await self._get_hwnd()
        self._frame = await frame
        await self._frame.attach(self, self.hwnd)
        _task_panes.add(self)
        
    async def _get_hwnd(self) -> int:
        """
            Should be implemented by derived classes
        """
        raise NotImplementedError()

    def on_visible(self, value):
        """ 
            Called when the visible state changes, `value` contains the new state.
            It is not necessary to override this to control pane visibility - the
            window will be shown/hidden automatically
        """
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
    def size(self) -> typing.Tuple[int, int]:
        """Returns a tuple (width, height)"""
        return self._frame.size

    @size.setter
    def size(self, value: typing.Tuple[int, int]):
        self._frame.size = value


def find_task_pane(title:str=None, workbook=None, window=None) -> CustomTaskPane:
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


def _try_create_from_qwidget(obj) -> typing.Awaitable[CustomTaskPane]:
    """
        If obj is a QWidget, returns an Awaitable[QtThreadTaskPane] else None.
        This is a convenience for Qt users to avoid needing to create a QtThreadTaskPane
        explicitly
    """
    try:
        # This will raise an ImportError if Qt has not been properly
        # setup by importing xloil.gui.pyqt5 (or pyside2). This check is
        # very cheap: _qtconfig imports nothing.
        from ._qtconfig import QT_IMPORT
        QWidget = QT_IMPORT("QtWidgets").QWidget
        
        if isinstance(obj, type) and issubclass(obj, QWidget) or isinstance(obj, QWidget):
            from ._qtgui import QtThreadTaskPane, Qt_thread
            return Qt_thread().submit_async(QtThreadTaskPane, obj)

    except ImportError:
        pass

    return obj


async def _attach_task_pane(
        gui: ExcelGUI,
        pane: CustomTaskPane,
        name: str, 
        window: ExcelWindow, 
        size: tuple, 
        visible: bool):
    
    # This is the implementation of ExcelGUI.attach_pane since the async
    # stuff and checking for Qt is easier on the python side

    if isinstance(window, str):
        window = ExcelWindow(window)
    
    if name is None:
        name = getattr(pane, "name", None)
        if name is None:
            raise NameError("Attach pane must be given a 'name' argument or the pane must "
                "contain a 'name' attribute")

    frame_future = gui._create_task_pane_frame(name, window)

    # A little convenience for Qt users to avoid needing to create a QtThreadTaskPane
    pane = _try_create_from_qwidget(pane)

    if inspect.isawaitable(pane):
        pane = await pane

    await pane._attach_frame(frame_future)
    
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
        import queue

        self._work_queue = queue.SimpleQueue()
        self._thread = threading.Thread(target=self._main_loop, name=name)
        self._broken = False
        self._thread.start()

        # PyBye is called before `threading` module teardown, whereas `atexit` comes later.
        # We definitely want threading available to shut down our threads.
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

    def submit(self, fn, *args, **kwargs) -> futures.Future:
        if self._broken:
            raise futures.BrokenExecutor(self._broken)

        future = futures.Future()
        work = concurrent.futures.thread._WorkItem(future, fn, args, kwargs)

        # In case we're called from our own thread - don't deadlock, just
        # execute the function right now.
        if threading.get_native_id() == self._thread.native_id:
            work.run()
        else:
            self._work_queue.put(work)

        return future

    async def submit_async(self, fn, *args, **kwargs):
        """
        Behaves as `submit` but wraps the result in an asyncio.Future so it
        can be awaited.
        """
        return await asyncio.wrap_future(self.submit(fn, *args, **kwargs))

    def shutdown(self, wait=True, cancel_futures=False):
        if not self._broken: # TODO: implement wait/cancel_futures support?
            self.submit(self._shutdown)

    def _main_loop(self):

        try:
            self._main()
        except Exception as e:
            xloil.log(f"{self._thread.name} failed: {e}", level='error')

        self._broken = True

    def _wrap(self, fn, sync=True, discard=False):
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
        
        if sync:
            async def wrapped(*args, **kwargs):
                return await self.submit_async(fn, *args, **kwargs)
        else:
            def wrapped(*args, **kwargs):
                return self.submit(fn, *args, **kwargs)

        return wrapped


# A metaclass which returns a metaclass. It could have been a function,
# just without the fun.
class _ConstructInExecutor:
    
    def __new__(cls_unused, executor):

        class _ConstructInExecutorImpl(type):

            def __call__(cls, *args, **kwargs):

                # Strictly only the __init__ method needs to be in the executor, but
                # this seems easier than replicating `type.__call__`
                if kwargs.pop("_no_executor", None) is True:
                    return type.__call__(cls, *args, **kwargs)
                else:
                    return executor.submit(
                        lambda: type.__call__(cls, *args, **kwargs)
                    ).result()
    
        return _ConstructInExecutorImpl

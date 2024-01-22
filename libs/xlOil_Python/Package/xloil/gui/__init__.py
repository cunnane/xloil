from xloil import TaskPaneFrame, ExcelGUI, ExcelWindow, Workbook

import concurrent.futures as futures
import concurrent.futures.thread
import xloil
import asyncio
import typing
import threading
import inspect
import sys

_TASK_PANES = set()

class CustomTaskPane:
    """
        Base class for custom task panes. xlOil provides two toolkit-specfic
        implementations: `xloil.gui.pyqt5.QtThreadTaskPane` (pyside is also 
        supported) and `xloil.gui.tkinter.TkThreadTaskPane`.
        
        Can be sub-classed to implement task panes with different GUI toolkits.

        Subclasses can implement functions to recieve events:

            on_visible(self, value):
                Called when the visible state changes, `value` contains the new state.
                It is not necessary to override this to control pane visibility - the
                window will be shown/hidden automatically

            on_docked(self):
                Called when the pane is docked to a new location or undocked
    """

    def __init__(self):
        self.hwnd = None

    async def _attach_frame_async(self, frame: typing.Awaitable[TaskPaneFrame]):
        """
            Attaches this *CustomTaskPane* to a *TaskPaneFrame*  causing it 
            to be resized/moved with the pane window.  Called by 
            `xloil.ExcelGUI.attach_frame` and generally should not need to be 
            called directly.
        """
        self.hwnd = await asyncio.wrap_future(self._get_hwnd())
        self._pane = await frame
        await self._pane.attach(self, self.hwnd)
        _TASK_PANES.add(self)

    def _attach_frame(self, frame: typing.Awaitable[TaskPaneFrame]):
        """
            Attaches this *CustomTaskPane* to a *TaskPaneFrame*  causing it 
            to be resized/moved with the pane window.  Called by 
            `xloil.ExcelGUI.attach_frame` and generally should not need to be 
            called directly.
        """
        self.hwnd = self._get_hwnd().result()
        self._pane = frame.result()
        self._pane.attach(self, self.hwnd).result()
        _TASK_PANES.add(self)
        
    def _get_hwnd(self) -> typing.Awaitable[int]:
        """
            Should be implemented by derived classes
        """
        raise NotImplementedError()

    def on_destroy(self):
        """ Called before the pane is destroyed to release any resources """
        # Release internal task pane pointer
        self._pane = None
        # Remove ourselves from pane lookup table
        _TASK_PANES.discard(self)

    @property
    def pane(self) -> TaskPaneFrame:
        """Returns the TaskPaneFrame: a reference to the window holding the python GUI"""
        return self._pane

    @property
    def visible(self) -> bool:
        """Returns True if the pane is currently shown"""
        return self._pane.visible

    @visible.setter
    def visible(self, value: bool):
        self._pane.visible = value

    @property
    def size(self) -> typing.Tuple[int, int]:
        """Returns a tuple (width, height)"""
        return self._pane.size

    @size.setter
    def size(self, value: typing.Tuple[int, int]):
        self._pane.size = value

    @property
    def position(self) -> str:
        """
        Returns the docking position: one of bottom, floating, left, right, top
        """
        return self._pane.position

def find_task_pane(title:str=None, workbook=None, window=None) -> CustomTaskPane:
    """
        Finds xlOil python task panes attached to the specified window, with the 
        given pane `title`. The primary use of this function is to look for an existing 
        task pane before creating a new one.

        Parameters
        ----------

        title:
            if `title` is specified, returns a (case-sensitive) match of a single 
            `xloil.CustomTaskPane` object or None if not found.  Otherwise returns a list 
            of `xloil.CustomTaskPane` objects.

        window: str or `xloil.ExcelWindow`
            The window title to be searched    

        workbook: str or `xloil.Workbook`:
            Task panes are linked to Excel's Window objects which can have a many-to-one
            relationship with workbooks. If a workbook is specified, all task panes 
            associated with that workbook will be searched.

    """

    if window is not None:
        if isinstance(window, str):
            window = ExcelWindow(window)
        hwnds = [window.hwnd]
    elif workbook is not None:
        if isinstance(workbook, str):
            workbook = Workbook(workbook)
        hwnds = [x.hwnd for x in workbook.windows]
    else:
        hwnds = [xloil.app().windows.active.hwnd]
 
    found = [x for x in _TASK_PANES if isinstance(x, CustomTaskPane) and x.pane.window.hwnd in hwnds]

    if title is None:
        return found
    else:
        return next((x for x in found if x.pane.title == title), None)


def _try_create_from_qwidget(obj) -> CustomTaskPane:
    """
        If obj is a QWidget, returns an Awaitable[QtThreadTaskPane] else None.
        This is a convenience for Qt users to avoid needing to create a QtThreadTaskPane
        explicitly
    """
    try:
        if 'qtpy' in sys.modules:
            from qtpy.QtWidgets import QWidget
            if (isinstance(obj, type) and issubclass(obj, QWidget)) or isinstance(obj, QWidget):
                from xloil.gui.qtpy import QtThreadTaskPane, Qt_thread
                return Qt_thread().submit(QtThreadTaskPane, obj).result() 
    except ImportError:
        pass

    return obj

def _try_create_from_wxframe(obj) -> CustomTaskPane:
    if 'wx' in sys.modules:
        import wx
        if (isinstance(obj, type) and issubclass(obj, wx.Frame)) or isinstance(obj, wx.Frame):
            from .wx import WxThreadTaskPane, wx_thread
            return wx_thread().submit(WxThreadTaskPane, obj).result()
    return obj

def _get_pane_name(pane):
    name = getattr(pane, "name", None)
    if name is None:
        raise NameError("Attach pane must be given a 'name' argument or the pane must "
            "contain a 'name' attribute")
    return name

async def _attach_task_pane_async(
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

    name = name or _get_pane_name(pane)
     
    frame_future = gui._create_task_pane_frame(name, window)

    # A little convenience for Qt/Wx users to avoid needing to create a QtThreadTaskPane
    pane = _try_create_from_qwidget(pane)
    pane = _try_create_from_wxframe(pane)

    if inspect.isawaitable(pane):
        pane = await pane

    await pane._attach_frame_async(frame_future)
    
    pane.visible = visible
    if size is not None:
        pane.size = size

    return pane


def _attach_task_pane(
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

    name = name or _get_pane_name(pane)

    frame_future = gui._create_task_pane_frame(name, window)

    # A little convenience for Qt users to avoid needing to create a QtThreadTaskPane
    pane = _try_create_from_qwidget(pane)
    pane = _try_create_from_wxframe(pane)

    pane._attach_frame(frame_future)
    
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
        self._ready = futures.Future()
        self._thread = threading.Thread(target=self._main_loop, name=name)
        self._thread.start()

        # PyBye is called before `threading` module teardown, whereas `atexit` comes later.
        # We definitely want threading available to shut down our threads.
        xloil.event.PyBye += self.shutdown

        #xloil.log.error(f"Did we make it here")

    def _make_ready(self):
        """
        Should be called by derive classes to signal that it is safe to add jobs to the queue
        """
        self._ready.set_result(True)

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
        """
        Schedules the callable, fn, to be executed as ``fn(*args, **kwargs)`` and returns 
        a ``Future`` object representing the execution of the callable.
        """

        if not self._ready.result():
            raise futures.BrokenExecutor()

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
        if self._ready.result(): # TODO: implement cancel_futures?
            self.submit(self._shutdown)
        if wait:
            self._thread.join()
            xloil.log.debug("Joined GUI thread '%s'", self._thread.name)

    def _main_loop(self):

        name = self._thread.name
        try:
            xloil.log(f"Starting executor '{name}'", level="info")
            self._main()
            xloil.log(f"Finalising executor '{name}'", level="info")
            
        except Exception as e:
            xloil.log(f"{name} failed: {e}", level='error')
            self._ready.set_result(False)
            raise

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


class _ConstructInExecutor(type):
    """
        Metaclass used by CustomTaskPane objects to ensure their ctor is run in 
        the thread associated with the GUI toolkit. The `executor` argument should 
        specify a concurrent.futures.Executor. If None is passed, the functionality of 
        this metaclass is disabled. The default is `executor=True` which means take
        the executor from a base class, or throw if none is found.
    """

    def __new__(cls, name, bases, namespace, executor:futures.Executor = True):
        return type.__new__(cls, name, bases, namespace)
        
    def __init__(cls, name, bases, namespace, executor:futures.Executor = True):
        # If default, try to fetch executor from base class
        if executor == True:
            cls._executor = next(b._executor for b in bases if type(b) is _ConstructInExecutor)
        else:
            cls._executor = executor
            
        return type.__init__(cls, name, bases, namespace)
        
    def __call__(cls, *args, **kwargs):
        # Strictly only the __init__ method needs to be in the executor, but
        # this seems easier than replicating `type.__call__`
        if cls._executor is None:
            return type.__call__(cls, *args, **kwargs)
        else:
            return cls._executor().submit(
                lambda: type.__call__(cls, *args, **kwargs)
            ).result()
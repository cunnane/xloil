import threading
import queue
from ._common import *
from .shadow_core import event
from .excelgui import CustomTaskPane
import importlib
import sys
import concurrent.futures as futures
import concurrent.futures.thread

def qt_import(sub, what):
    """
        Helper function to import Q objects from PyQt or PySide depending
        on which framework had already been imported
    """
    if 'PyQt5' in sys.modules:
        top = 'PyQt5'
    elif 'PySide2' in sys.modules:
        top = 'PySide2'
    else:
        raise ImportError("Import PyQt or PySide before invoking this function")

    if isinstance(what, str):
        mod = __import__(top + '.' + sub, globals(), locals(), [what], 0)
        return getattr(mod, what)
    else:
        mod = __import__(top + '.' + sub, globals(), locals(), what, 0)
        return [getattr(mod, x) for x in what]
     

def _create_Qt_app():

    # For some reason, my version of PyQt doesn't read the platform plugin
    # path env var, so I need to explicitly pass it to the QApplication ctor
    import os
    ppp = os.environ.get('QT_QPA_PLATFORM_PLUGIN_PATH')

    QApplication = qt_import('QtWidgets', 'QApplication')

    app = QApplication([] if ppp is None else ['','-platformpluginpath', ppp])

    log(f"Started Qt on thread {threading.get_native_id()}" +
        f"with libpaths={app.libraryPaths()}", level="info")

    return app

class QtExecutor(futures.Executor):

    def __init__(self):
        self._work_queue = queue.SimpleQueue()
        self._thread = threading.Thread(target=self._main_loop, name="QtGuiThread")
        self._broken = False
        self._work_signal = None
        self._thread.start()

    def submit(self, fn, *args, **kwargs):
        if self._broken:
            raise futures.BrokenExecutor(self._broken)

        f = futures.Future()
        w = concurrent.futures.thread._WorkItem(f, fn, args, kwargs)

        self._work_queue.put(w)
        if self._work_signal is not None:
            self._work_signal.timeout.emit()
        return f

    def shutdown(self, wait=True, cancel_futures=False):
        if not self._broken:
            self.submit(self.app.quit)

    def _do_work(self):
        try:
            while True:
                work_item = self._work_queue.get_nowait()
                if work_item is not None:
                    work_item.run()
                    del work_item
        except queue.Empty:
            return
            
    def _main_loop(self):

        try:
            self.app = _create_Qt_app()

            QTimer = qt_import('QtCore', 'QTimer')

            semaphore = QTimer()
            semaphore.timeout.connect(self._do_work)
            self._work_signal = semaphore

            # Trigger timer to run any pending queue items now
            semaphore.timeout.emit() 

            # Thread main loop, run until quit
            self.app.exec()

            # Thread cleanup
            self.app = None
            self._enqueued = None
            self._broken = True

        except Exception as e:
            self._broken = True
            log(f"QtThread failed: {e}", level='error')


_Qt_thread = None

def Qt_thread():
    """
        Since all Qt GUI interactions (except signals) must take place on the 
        thread that the QApplication object was created on, we have a dedicated
        thread with a work queue.
    """

    global _Qt_thread

    if _Qt_thread is None:
        _Qt_thread = QtExecutor()
        # PyBye is called before `threading` module teardown, whereas `atexit` comes later
        event.PyBye += _Qt_thread.shutdown
        # Send this blocking no-op to ensure QApplication is created on our thread
        # before we proceed, otherwise Qt may try to create one elsewhere
        _Qt_thread.submit(lambda: 0).result()

    return _Qt_thread


class QtThreadTaskPane(CustomTaskPane):

    def __init__(self, pane, draw_widget):
        """
        Wraps a QWidget to create a CustomTaskPane object. The `draw_widget` function
        is executed on the QtThread and is expected to return a QWidget object.
        """
        super().__init__(pane)

        self.widget = Qt_thread().submit(draw_widget).result() # Blocks
        Qt_thread().submit(lambda: self._reparent_widget(self.widget, self.pane.parent_hwnd))

    def on_size(self, width, height):
        Qt_thread().submit(lambda: self.widget.resize(width, height))
             
    def on_visible(self, c):
        Qt_thread().submit(lambda: self.widget.show() if c else self.widget.hide())

    def on_destroy(self):
        Qt_thread().submit(lambda: self.widget.destroy())
        super().on_destroy()

    def _reparent_widget(self, widget, hwnd):
        QWindow = qt_import('QtGui', 'QWindow')
        # windowHandle does not exist before show
        widget.show()
        nativeWindow = QWindow.fromWinId(hwnd)
        widget.windowHandle().setParent(nativeWindow)
        widget.update()
        widget.move(0, 0)


def _try_create_qt_pane(obj):
    try:
        QWidget = qt_import('QtWidgets', 'QWidget')
        if issubclass(obj, QWidget):
            return lambda pane: QtThreadTaskPane(pane, obj)
    except ImportError:
        pass

    return None

"""
    Do not import this module directly
"""


import sys
import concurrent.futures as futures
import concurrent.futures.thread
from xloil.gui import CustomTaskPane, _GuiExecutor
from xloil import log
from ._qtconfig import QT_IMPORT
import threading
import xloil

def _create_Qt_app():

    QApplication = QT_IMPORT("QtWidgets").QApplication

    # Qt seems to really battle with reading environment variables.  We must 
    # read the variable ourselves, then pass it as an argument. It's unclear what
    # alchemy is required to make Qt do this seemingly simple thing itself.
    import os
    ppp = os.getenv('QT_QPA_PLATFORM_PLUGIN_PATH', None)
    app = QApplication([] if ppp is None else ['','-platformpluginpath', ppp])

    log(f"Started Qt on thread {threading.get_native_id()}" +
        f"with libpaths={app.libraryPaths()}", level="error")

    return app

def _reparent_widget(widget, hwnd):

    QWindow = QT_IMPORT("QtGui").QWindow

    # windowHandle does not exist before show
    widget.show()
    nativeWindow = QWindow.fromWinId(hwnd)
    widget.windowHandle().setParent(nativeWindow)
    widget.update()
    widget.move(0, 0)


class QtExecutor(_GuiExecutor):

    def __init__(self):
        self._work_signal = None
        self._app = None
        super().__init__("QtGuiThread")

    @property
    def app(self):
        return self._app

    def submit(self, fn, *args, **kwargs):
        future = super().submit(fn, *args, **kwargs)
        if self._work_signal is not None:
            self._work_signal.timeout.emit()
        return future

    def _main(self):
        self._app = _create_Qt_app()

        QTimer = QT_IMPORT("QtCore").QTimer

        semaphore = QTimer()
        semaphore.timeout.connect(self._do_work)
        self._work_signal = semaphore

        # Run any pending queue items now
        self._do_work()

        # Thread main loop, run until quit
        self._app.exec()

        # Thread cleanup
        self._app = None

    def _shutdown(self):
        self._app.quit()


_Qt_thread = None

def Qt_thread() -> futures.Executor:
    """
        All Qt GUI interactions (except signals) must take place on the thread on which 
        the *QApplication* object was created.  This object is a *concurrent.futures.Executor* 
        which creates the *QApplication* object and can run commands on a dedicated Qt thread.  
        
        **All QT interaction must take place via this thread**.

        Examples
        --------
            
        ::
            
            future = Qt_thread().submit(my_func, my_args)
            future.result() # blocks

    """

    global _Qt_thread

    if _Qt_thread is None:
        _Qt_thread = QtExecutor()
        # Send this blocking no-op to ensure QApplication is created on our thread now
        _Qt_thread.submit(lambda: 0).result()

    return _Qt_thread


class QtThreadTaskPane(CustomTaskPane):
    """
        Wraps a Qt QWidget to create a CustomTaskPane object. 
    """

    def __init__(self, pane, draw_widget):
        """
        Wraps a QWidget to create a CustomTaskPane object. The ``draw_widget`` function
        is executed on the `xloil.gui.Qt_thread` and is expected to return a *QWidget* object.
        """
        super().__init__(pane)

        def draw_it(hwnd):
            widget = draw_widget()
            _reparent_widget(widget, hwnd)
            return widget
        self.widget = Qt_thread().submit(draw_it, self.pane.parent_hwnd).result() # Blocks

    def on_size(self, width, height):
        Qt_thread().submit(lambda: self.widget.resize(width, height))
             
    def on_visible(self, c):
        Qt_thread().submit(lambda: self.widget.show() if c else self.widget.hide())

    def on_destroy(self):
        Qt_thread().submit(lambda: self.widget.destroy())
        super().on_destroy()


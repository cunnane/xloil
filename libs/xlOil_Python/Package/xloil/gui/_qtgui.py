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

    log(f"Started Qt on thread {threading.get_native_id()} " +
        f"with libpaths={app.libraryPaths()}", level="info")

    return app


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

def Qt_thread(fn=None, discard=False) -> futures.Executor:
    """
        All Qt GUI interactions (except signals) must take place on the thread on which 
        the *QApplication* object was created.  This object returns a *concurrent.futures.Executor* 
        which creates the *QApplication* object and can run commands on a dedicated Qt thread. 
        It can also be used a decorator.
        
        **All QT interaction must take place via this thread**.

        Examples
        --------
            
        ::
            
            future = Qt_thread().submit(my_func, my_args)
            future.result() # blocks

            @Qt_thread(discard=True)
            def myfunc():
                # This is run on the Qt thread and returns a *future* to the result.
                # By specifying `discard=True` we tell xlOil that we're not going to
                # keep track of that future and so it should log any exceptions.
                ... 

    """

    global _Qt_thread

    if _Qt_thread is None:
        _Qt_thread = QtExecutor()
        # Send this blocking no-op to ensure QApplication is created on our thread now
        _Qt_thread.submit(lambda: 0).result()

    return _Qt_thread if fn is None else _Qt_thread._wrap(fn, discard)

Qt_thread()

class QtThreadTaskPane(CustomTaskPane):
    """
        Wraps a Qt QWidget to create a CustomTaskPane object. 
    """

    def __init__(self, widget=None):
        """
        Wraps a QWidget to create a CustomTaskPane object. The ``draw_widget`` function
        is executed on the `xloil.gui.Qt_thread` and is expected to return a *QWidget* object.
        """
        super().__init__()

        def _get_hwnd(obj):
            obj.show() # window handle does not exist before show
            return int(obj.winId())

        def draw_it():
            QWidget = QT_IMPORT("QtWidgets").QWidget

            if isinstance(widget, QWidget):
                obj = widget
            elif widget is not None:
                obj = widget()
            else:
                self.draw()

            # Need to make the Qt window frameless using Qt's API. When we attach
            # to the TaskPaneFrame, the attached window is turned into a 
            # frameless child. If Qt is not informed, its geometry manager gets
            # confused and will core dump if the pane is made too small.
            qt = QT_IMPORT("QtCore").Qt
            obj.setWindowFlags(qt.FramelessWindowHint)

            return obj, _get_hwnd(obj)

        self.contents, self._hwnd = Qt_thread().submit(draw_it).result() # Blocks

    def hwnd(self):
        return self._hwnd

    def on_visible(self, c):
        Qt_thread().submit(lambda: self.contents.show() if c else self.contents.hide())

    def on_destroy(self):
        Qt_thread().submit(lambda: self.contents.destroy())
        super().on_destroy()


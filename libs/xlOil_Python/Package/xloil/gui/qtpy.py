"""
    You must import this module before any other mention of Qt or Pyside: this allows xlOil to 
    create a thread to manage the Qt GUI and the Qt app object.  *All* interaction with the 
    *Qt* must be done on that thread or crashes will ensue.  Use `Qt_thread.submit(...)`
    or the `@Qt_thread` to ensure functions are run on the thread.
"""
import sys
import concurrent.futures as futures
import concurrent.futures.thread
from xloil.gui import CustomTaskPane, _GuiExecutor, _ConstructInExecutor
from xloil import log
import xloil
from xloil._core import XLOIL_EMBEDDED
import threading

def _create_Qt_app():

    from qtpy import QtCore
    def qt_msg_handler(msg_type, msg_log_context, msg_string):

        level = 'info'
        if msg_type == QtCore.QtDebugMsg:
            level = 'debug'
        elif msg_type == QtCore.QtInfoMsg:
            level = 'info'
        elif msg_type == QtCore.QtWarningMsg:
            level = 'warn'
        elif msg_type == QtCore.QtCriticalMsg or msg_type == QtCore.QtFatalMsg:
            level = 'error'

        # Qt raises this on shutdown. No idea why, but don't want it to trigger
        # a log window popup!
        if msg_string.startswith('QWindowsBackingStore::flush: GetDC'):
            level = 'debug'

        log(msg_string, level=level)

    QtCore.qInstallMessageHandler(qt_msg_handler)

    log.info(f"Starting Qt on thread {threading.get_native_id()}")
    app = QApplication([])
    return app

class QtExecutor(_GuiExecutor):

    def __init__(self):
        self._work_signal = None
        self._app = None
        super().__init__("QtGuiThread")

    @property
    def app(self):
        """
            A reference to the singleton *QApplication* object 
        """
        return self._app

    def submit(self, fn, *args, **kwargs):
        future = super().submit(fn, *args, **kwargs)
        if self._work_signal is not None:
            self._work_signal.timeout.emit()
        return future

    def _main(self):
        self._app = _create_Qt_app()

        from qtpy.QtCore import QTimer
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

def Qt_thread(fn=None, discard=False) -> QtExecutor:
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
        # Send this blocking job to ensure QApplication is created now
        _Qt_thread.submit(lambda: 0).result()

    return _Qt_thread if fn is None else _Qt_thread._wrap(fn, discard)

# For safety, initialise Qt when we this module is imported
if XLOIL_EMBEDDED:
    Qt_thread()

class QtThreadTaskPane(CustomTaskPane, metaclass=_ConstructInExecutor, executor=Qt_thread):
    """
        Wraps a Qt *QWidget* to create a `CustomTaskPane` object. The optional `widget` argument 
        must be a type deriving from *QWidget* or an instance of such a type (a lambda which
        returns a *QWidget* is also acceptable).
    """

    def __init__(self, widget=None):
        from qtpy.QtWidgets import QWidget
        if isinstance(widget, QWidget):
            self._widget = widget
        elif widget is not None:
            self._widget = widget()
        else:
            self._widget = QWidget()
    
    @property
    def widget(self):
        """
            This returns the *QWidget* which is root of the the pane's contents.
            If the class was constructed from a *QWidget*, this is that widget.
        """
        return self._widget

    def _get_hwnd(self):
        def prepare(widget):
            # Need to make the Qt window frameless using Qt's API. When we attach
            # to the TaskPaneFrame, the attached window is turned into a frameless
            # child. If Qt is not informed, its geometry manager gets confused
            # and will core dump if the pane is made too small.
            from qtpy.QtCore import Qt
            widget.setWindowFlags(Qt.FramelessWindowHint)

            widget.show() # window handle does not exist before show

            return int(widget.winId())

        return Qt_thread().submit(prepare, self._widget)

    def on_destroy(self):
        # Call super to detach the TaskPaneFrame for a cleaner shutdown
        super().on_destroy() 
        Qt_thread().submit(lambda: self._widget.destroy())

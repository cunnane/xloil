import threading
import queue
from ._common import *
from .shadow_core import event
from .excelgui import CustomTaskPane

class _QtThread:

    def __init__(self):
        self._thread = None
        self._app = None
        self._enqueued = None

    def start(self):
        if self._thread is not None:
            return
        self._thread = threading.Thread(target=self._main_loop, name="QtGuiThread")
        self._queue = queue.Queue()
        self._results = queue.Queue()
        self._thread.start()
        # PyBye is called before threading module teardown, whereas `atexit` comes later
        event.PyBye += self.stop

    def stop(self):
        if self.ready:
            self._queue.put((False, self.app.quit))
            self._enqueued.timeout.emit()
        
    def run(self, cmd):
        self._queue.put((True, cmd))
        if self.ready:
            self._enqueued.timeout.emit()

        result = self._results.get() # Blocks
        if isinstance(result, Exception):
            raise result
        return result

    def send(self, cmd):
        if self.stopped:
            raise RuntimeError("Qt GUI Thread has stopped unexpectedly")
        self._queue.put((False, cmd))
        if self.ready:
            self._enqueued.timeout.emit()

    @property
    def ready(self):
        return not self.stopped and self._enqueued is not None
    
    @property
    def stopped(self):
        return not self._thread or not self._thread.is_alive()
        
    @property
    def app(self):
        return self._app
        
    def _main_loop(self):
 
        # For some reason, my version of PyQt doesn't read the platform plugin
        # path env var, so I need to explicitly pass it to the QApplication ctor
        try:
            import os
            ppp = os.environ['QT_QPA_PLATFORM_PLUGIN_PATH']
            from PyQt5.QtWidgets import QApplication
            from PyQt5.QtCore import QTimer
        
            self._app = QApplication(['','-platformpluginpath', ppp])

            log(f"Started Qt on thread {threading.get_native_id()} with libpaths={self._app.libraryPaths()}", level="info")

            def check_queue():
                try:
                    while True:
                        keep, item = self._queue.get_nowait()
                        try:
                            result = item()
                        except Exception as e:
                            result = e
                        if keep:
                            self._results.put(result)
                        self._queue.task_done()
                except queue.Empty:
                    return
            
            timer = QTimer() # Is there a better signal than this timer?
            timer.timeout.connect(check_queue)
            self._enqueued = timer

            # Trigger timer to run any pending queue items now
            timer.timeout.emit() 

            # Thread main loop, run until quit
            self._app.exec()
            # Thread cleanup
            self._app = None
            self._enqueued = None
        except Exception as e:
            log(f"QtThread failed: {e}", level='error')



QtThread = _QtThread() 

class QtThreadTaskPane(CustomTaskPane):

    def __init__(self, pane, draw_widget):
        super().__init__(pane)

        # Send this blocking no-op to ensure our QApplication is created first as importing 
        # PyQt5.QtWidgets will create one if it does not already exist.
        QtThread.start()
        QtThread.run(lambda: 0)

        self.widget = QtThread.run(draw_widget)
        QtThread.run(lambda: self._reparent_widget(self.widget, self.pane.parent_hwnd))

    def on_size(self, width, height):
        QtThread.send(lambda: self.widget.resize(width, height))
             
    def on_visible(self, c):
        QtThread.send(lambda: self.widget.show() if c else self.widget.hide())

    def on_destroy(self):
        QtThread.send(lambda: self.widget.destroy())
        super().on_destroy()

    def _reparent_widget(self, widget, hwnd):
        from PyQt5.QtGui import QWindow
        # windowHandle does not exist before show
        widget.show() 
        nativeWindow = QWindow.fromWinId(hwnd)
        widget.windowHandle().setParent(nativeWindow)
        widget.update()
        widget.move(0, 0)

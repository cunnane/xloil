import threading
import queue
import xloil

class _QtThread:

    def __init__(self):
        self._thread = threading.Thread(target=self._main_loop, name="QtGuiThread")
        self._app = None
        self._enqueued = None
        self._queue = queue.Queue()
        self._results = queue.Queue()
        self._thread.start()

            
    def join(self):
        self._thread.join()
        
    def stop(self):
        xloil.log("Thread stopped", level="error")
        if self.app:
            self._queue.put(self.app.quit)
            self._enqueued.timeout.emit()
        
    def send(self, cmd):
        if not self.ready:
            raise RuntimeError()
        self._queue.put(cmd)
        self._enqueued.timeout.emit()
        self._queue.join()
        return self._results.get()
    
    @property
    def ready(self):
        return self._enqueued is not None
    
    @property
    def stopped(self):
        return not self._thread.is_alive()
        
    @property
    def app(self):
        return self._app
        
    def _main_loop(self):
 
        #from PyQt5.QtCore import QCoreApplication

        #app = QCoreApplication([])

        import os
        ppp = os.environ['QT_QPA_PLATFORM_PLUGIN_PATH']
        #xloil.log(f"Qt with ppp={ppp}", level="error")
        #app.addLibraryPath(ppp)

        # TODO: another version of pyqt?
        from PyQt5.QtWidgets import QApplication
        from PyQt5.QtCore import QTimer
        
        self._app = QApplication(['','-platformpluginpath', ppp])

        xloil.log(f"Started Qt with libpaths={self._app.libraryPaths()}", level="error")
        def check_queue():
            try:
                while True:
                    item = self._queue.get_nowait()
                    self._results.put(item())
                    self._queue.task_done()
            except queue.Empty:
                return
            
        timer = QTimer()
        timer.timeout.connect(check_queue)
        self._enqueued = timer
        self._app.exec()
        self._app = None
        self._enqueued = None


QtThread = _QtThread()

# PyBye is called before threading module teardown, whereas `atexit` is not.
_stopper = QtThread.stop
xloil.event.PyBye += _stopper
import xloil.gui.qtpy
import xloil

def create_qtconsole_inprocess():
    """
    Opens *qtconsole* using an inprocess ipython kernel. 
    Warning: using qtconsole can trigger a 'feature' in PyQt which causes it
    to dump core on application exit when it tries to access the python 
    interpreter after xloil has unloaded.
    """
    from qtconsole.rich_jupyter_widget import RichJupyterWidget
    from qtconsole.inprocess import QtInProcessKernelManager, QtInProcessRichJupyterWidget
    from xloil.gui.qtpy import Qt_thread
    from qtpy.QtWidgets import QMainWindow
    import asyncio
    import sys
    import threading

    xloil.log.debug("Starting QtConsole on thread %d", threading.get_native_id())

    # Need to set these dummy streams as the inprocess kernel calls
    # flush on stdout and stderr 
    class DummyStream:
        def write(self, string):
            xloil.log.debug(string)

        def flush(self):
            pass

        def readline(self):
            return None

    sys.stdout = DummyStream()
    sys.stderr = DummyStream()

    def make_jupyter_widget():
        kernel_manager = QtInProcessKernelManager()
        kernel_manager.start_kernel(show_banner=False)
        kernel = kernel_manager.kernel
        kernel.gui = 'qt4'

        kernel_client = kernel_manager.client()
        kernel_client.start_channels()

        ipython_widget = QtInProcessRichJupyterWidget()
        ipython_widget.kernel_manager = kernel_manager
        ipython_widget.kernel_client = kernel_client
        ipython_widget.setObjectName("QtConsole")
        return ipython_widget

    class MainWindow(QMainWindow):

        def __init__(self):
            super().__init__()
            self.jupyter_widget = make_jupyter_widget()
            self.setCentralWidget(self.jupyter_widget)

        def closeEvent(self, event):
            try:
                self.shutdown_kernel()
                ...
            except Exception as e:
                # Throwing from this function causes Qt to dump core
                xloil.log_except("Failed shutting down QtConsole")

        def shutdown_kernel(self):
            xloil.log.debug("Shutting QtConsole on thread %d",
                            threading.get_native_id())
            # We force the kernels atexit routine to run now, otherwise it 
            # shuts down the SQLite server used to keep history on the wrong 
            # thread and generates "SQLite objects created in a thread can
            # only be used in that same thread."
            self.jupyter_widget.kernel_manager.kernel.shell._atexit_once()
            self.jupyter_widget.kernel_client.stop_channels()
            self.jupyter_widget.kernel_manager.shutdown_kernel()
            xloil.log.debug("Finished Shutting QtConsole on thread")
    
    from qtpy.QtCore import Qt
    console = MainWindow()

    # Not doing this seems to make Qt more likely to dump core on exit
    console.setAttribute(Qt.WA_DeleteOnClose)
    return console

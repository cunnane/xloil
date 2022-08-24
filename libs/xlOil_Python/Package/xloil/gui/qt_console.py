import xloil.gui.qtpy

def create_qtconsole_inprocess():
    """
    Opens *qtconsole* using an inprocess ipython kernel. Currently experimental.
    """
    from qtconsole.rich_jupyter_widget import RichJupyterWidget
    from qtconsole.inprocess import QtInProcessKernelManager
    import asyncio
    import sys

    asyncio.set_event_loop(asyncio.new_event_loop())

    # Need to set these dummy streams as the inprocess kernel calls
    # flush on stdout and stderr 
    class DummyStream:
        def write(self, string):
            xloil.log.trace(string)

        def flush(self):
            pass

        def readline(self):
            return None

    sys.stdout = DummyStream()
    sys.stderr = DummyStream()

    kernel_manager = QtInProcessKernelManager()
    kernel_manager.start_kernel(show_banner=False)
    kernel = kernel_manager.kernel
    kernel.gui = 'qt4'

    kernel_client = kernel_manager.client()
    kernel_client.start_channels()

    ipython_widget = RichJupyterWidget()
    ipython_widget.kernel_manager = kernel_manager
    ipython_widget.kernel_client = kernel_client
    return ipython_widget

import xloil.gui.qtpy
import xloil

def create_qt_ipython_console(run_on_excel_thread=True):
    """
    Opens *qtconsole* using an inprocess ipython kernel.
    """
    from qtconsole.rich_jupyter_widget import RichJupyterWidget
    from qtconsole.inprocess import QtInProcessKernelManager
    import asyncio
    import sys
    from xloil.ipython_kernel import MainThreadInProcessKernel, _DummyStream

    class OurQtInProcessKernelManager(QtInProcessKernelManager):
        def start_kernel(self, **kwds):
            self.kernel = MainThreadInProcessKernel(parent=self, session=self.session)

    asyncio.set_event_loop(asyncio.new_event_loop())

    sys.stdout = _DummyStream()
    sys.stderr = _DummyStream()

    kernel_manager = OurQtInProcessKernelManager() if run_on_excel_thread else QtInProcessKernelManager()
    kernel_manager.start_kernel(show_banner=False)

    kernel = kernel_manager.kernel
    kernel.gui = 'qt4'

    kernel_client = kernel_manager.client()
    kernel_client.start_channels()

    ipython_widget = RichJupyterWidget()
    ipython_widget.kernel_manager = kernel_manager
    ipython_widget.kernel_client = kernel_client
    return ipython_widget

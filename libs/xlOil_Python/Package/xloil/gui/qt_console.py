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
    from ipykernel.inprocess.ipkernel import InProcessKernel

    class OurInProcessKernel(InProcessKernel):
        async def execute_request(self, stream, ident, parent):
            try:
                run_it = super().execute_request
                return await asyncio.ensure_future(
                    xloil.excel_callback(
                        lambda: asyncio.get_event_loop().run_until_complete(
                            run_it(stream, ident, parent))))
            except e:
                xloil.log.error(str(e))

    class OurQtInProcessKernelManager(QtInProcessKernelManager):
        def start_kernel(self, **kwds):
            self.kernel = OurInProcessKernel(parent=self, session=self.session)

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

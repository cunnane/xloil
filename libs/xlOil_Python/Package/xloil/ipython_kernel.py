
import asyncio
import sys
import xloil

from .logging import get_logging_handler
from ipykernel.inprocess.ipkernel import InProcessKernel
from background_zmq_ipython import IPythonBackgroundKernelWrapper


class _DummyStream:
    """
    Dummy stream which can be used to set stdout/stderr as the IPython kernel base 
    tries to flush those streams and will complain if they don't exist
    """
    def write(self, string):
        xloil.log.trace(string)

    def flush(self):
        pass

    def readline(self):
        return None


class MainThreadInProcessKernel(InProcessKernel):
    """
    Extends IPython's InProcessKernel but runs all commands on Excel's main thread
    """
    async def execute_request(self, stream, ident, parent):
        try:
            run_it = super().execute_request
            return await asyncio.ensure_future(
                xloil.excel_callback(
                    lambda: asyncio.get_event_loop().run_until_complete(
                        run_it(stream, ident, parent))))
        except e:
            xloil.log.error(str(e))



class _OurIPythonBackgroundKernelWrapper(IPythonBackgroundKernelWrapper):

    def _create_kernel(self):

        super()._create_kernel()

        from jupyter_client.utils import run_sync
        real_do_execute = run_sync(self._kernel.do_execute)

        async def patched_do_execute(self, *args, **kwargs):
            def run():
                return real_do_execute(*args, **kwargs)
            return await asyncio.ensure_future(xloil.excel_callback(run))

        #def patched_do_execute(*args, **kwargs):
        #    return xloil.excel_callback(lambda: real_do_execute(*args, **kwargs)).result()

        self._kernel.do_execute = patched_do_execute
     
    def _start_kernel(self):
        super()._start_kernel()
        xloil.log.info(f"Main thread kernel listening on {self.connection_filename}: {str(self._connection_info)}")


_main_thread_kernel = None

def start_zmq_main_thread_kernel():
    """
    Starts an IPython kernel which runs commands on the main thread and can be connected
    to in the usual way via ZMQ.
    """
    logger = logging.Logger(name="xloil")
    log_handler = get_logging_handler()
    log_handler.setFormatter(logging.Formatter())
    logger.addHandler(log_handler)

    sys.stdout = _DummyStream()
    sys.stderr = _DummyStream()

    manager = _OurIPythonBackgroundKernelWrapper(
        logger=logger, banner="xlOil Kernel: Commands are run on Excel's main thread")
    manager.start()
    
    _main_thread_kernel = manager

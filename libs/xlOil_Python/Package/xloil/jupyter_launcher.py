
import os
import asyncio
import xloil
import sys

from jupyter_client.manager import KernelManager
from notebook.services.kernels.kernelmanager import MappingKernelManager
from asyncio.futures import Future
from jupyter_client.ioloop import IOLoopKernelManager

def _launch_jupyter(notebook_path:str, default_kernel_name:str):

    from subprocess import Popen
    from xloil._core import xloil_bin_path()

    xloil_path = xloil_bin_path()
    if not xloil_path in os.environ["PATH"]:
        os.environ["PATH"] += os.pathsep + xloil_path
    
    # The jupyter runner is installed to the Scripts directory which may not
    # be on the path
    scripts_dir = os.path.join(sys.prefix, "Scripts")
    if not scripts_dir in os.environ["PATH"]:
        os.environ["PATH"] += os.pathsep + scripts_dir

    process = Popen(['cmd.exe', '/k',
        "jupyter", "notebook", notebook_path,
          "--no-browser",
          "--NotebookApp.kernel_manager_class=xloil.jupyter_launcher.XlOilMultiKernelManager",
          f"--MultiKernelManager.default_kernel_name={default_kernel_name}",
          "--Session.key='b\"\"'"]) # Is this required?

    return process

def _change_notebook_kernel(filename: str, kernel_name:str):
    """
    Sets the kernel for the given notebook. The notebook is created if it
    does not exist
    """
    import nbformat as nbf
    
    if os.path.exists(filename):
        with open(filename, 'r') as f:
            notebook = nbf.read(f, nbf.NO_CONVERT)

    else:
        notebook = nbf.v4.new_notebook()
      
    notebook['metadata']['kernel_info'] = { 'name': kernel_name }
    with open(filename, 'w') as f:
        nbf.write(notebook, f)

# TODO: shutdown atexit
_attached_jupyter = None

async def open_attached_notebook():
    """
    Opens the jupyter notebook "attached" to the active workbook, that is,
    the one in the same directory with matching filename stem. Creates the
    file if it does not exist.
    """
    global _attached_jupyter

    active_workbook = await asyncio.ensure_future(
        xloil.excel_callback(lambda: xloil.active_workbook().FullName))

    if not any(active_workbook):
        raise FileNotFoundError("Need to save workbook so it has a path")

    notebook_path = os.path.splitext(active_workbook)[0] + ".ipynb"

    # Start our inprocess kernel and get the connection file name, which we use as the
    # kernel name. We assume the connection file is in the usual jupyter directory
    from xloil.inprocess_kernel import start_main_thread_zmq_kernel
    kernel_manager = start_main_thread_zmq_kernel()
    kernel_name = os.path.splitext(os.path.basename(kernel_manager.connection_filename))[0]
    xloil.log.info(f"Kernel name is {kernel_name}")
    _change_notebook_kernel(notebook_path, kernel_name)

    # Check if we started jupyter
    if _attached_jupyter is None:
        _attached_jupyter = _launch_jupyter(notebook_path, kernel_name)
    else:
        # Check if the process is still running
        try:
            _attached_jupyter.wait(timeout=0.001)
            _attached_jupyter = _launch_jupyter(notebook_path, kernel_name)
        except TimeoutExpired:
            # Need to use REST API to make jupyter open our notebook
            raise NotImplementedError()


class XlOilMultiKernelManager(MappingKernelManager):
    """
    Extend MappingKernelManager, which is the default jupyter multi-kernel manager, so that
    requests for our specially named kernel or default kernel will be linked to a kernel 
    already running inside the Excel process
    """
    def pre_start_kernel(self, kernel_name:str, kwargs):

        if kernel_name is None or kernel_name.startswith("xloil"):

            # Reimplement this check in the super-class
            kernel_id = kwargs.pop("kernel_id", self.new_kernel_id(**kwargs))
            if kernel_id in self:
                raise DuplicateKernelError("Kernel already exists: %s" % kernel_id)

            # The default kernel name should be set to our xloil kernel
            if kernel_name is None:
                kernel_name = self.default_kernel_name

            from jupyter_core.paths import jupyter_runtime_dir
            connection_filename = os.path.join(jupyter_runtime_dir(), kernel_name + ".json")

            kernel_manager = _ExistingKernelManager(connection_filename)
            xloil.log.info(f"XlOilMultiKernelManager: found kernel {connection_filename}")
            return kernel_manager, kernel_name, kernel_id

        return super().pre_start_kernel(kernel_name, kwargs)


class _ExistingKernelManager(IOLoopKernelManager):
    """
    A manager for an existing kernel specified by the connection file
    """

    def __init__(self, connection_file:str):
        self.load_connection_file(connection_file)
        self._ready = Future()
        self._ready.set_result(None)

    # --------------------------------------------------------------------------
    # Kernel management methods
    # --------------------------------------------------------------------------

    def start_kernel(self, *args, **kwargs):
        ...

    def shutdown_kernel(self, *args, **kwargs):
        ...

    def restart_kernel(self, now=False, **kwds):
        raise NotImplementedError("Cannot restart existing kernel.")

    @property
    def has_kernel(self):
        return True

    def interrupt_kernel(self):
        # Surely we can?
        raise NotImplementedError("Cannot interrupt existing kernel.")

    def signal_kernel(self, signum):
        raise NotImplementedError("Cannot signal existing kernel.")

    def is_alive(self):
        return True
    
    @property
    def ready(self):
        return self._ready

# Required?
#KernelManagerABC.register(InProcessKernelManager)
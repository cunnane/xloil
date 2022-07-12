import xloil
import concurrent.futures as futures
from xloil.gui import _GuiExecutor


class TkExecutor(_GuiExecutor):

    def __init__(self):
        self._root = None
        super().__init__("TkGuiThread")

    @property
    def root(self):
        return self._root

    def _do_work(self):
        super()._do_work(self)
        self._root.after(200, self._do_work)

    def _main(self):

        import tkinter as tk

        self._root = tk.Tk(baseName="xlOil")
        self._root.withdraw()
            
        # Run any pending queue items now
        self._do_work()
 
        # Thread main loop, run until quit
        self._root.mainloop()

        # Thread cleanup
        # Avoid Tcl_AsyncDelete: async handler deleted by the wrong thread
        self._root = None
        import gc
        gc.collect()

    def _shutdown(self):
        self._root.destroy()


_Tk_thread = None

def Tk_thread() -> futures.Executor:
    """
        All Tk GUI interactions must take place on the thread on which the root object 
        was created.  This object is a  *concurrent.futures.Executor* which creates the  
        root object and can run commands on a dedicated Tk thread.  
        
        **All Tk interaction must take place via this thread**.

        Examples
        --------
            
        ::
            
            future = Tk_thread().submit(my_func, my_args)
            future.result() # blocks

    """

    global _Tk_thread

    if _Tk_thread is None:
        _Tk_thread = TkExecutor()
        # Send this blocking no-op to ensure Tk is created on our thread now
        _Tk_thread.submit(lambda: 0).result()

    return _Tk_thread

# Create thread on import
Tk_thread()
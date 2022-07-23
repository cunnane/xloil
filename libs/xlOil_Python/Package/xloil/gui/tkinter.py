"""
    You must import this module before any other mention of `tkinter`: this allows xlOil to 
    create a thread to manage the Tk GUI and the Tk root object.  *All* interaction with the 
    *Tk* must be done on that thread or crashes will ensue.  Use `Tk_thread.submit(...)`
    or the `@Tk_thread` to ensure functions are run on the thread.
"""

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
        super()._do_work()
        self._root.after(200, self._do_work)

    def _main(self):

        import tkinter as tk

        self._root = tk.Tk(baseName="xlOil")
        self._root.withdraw()
            
        # Run any pending queue items now
        self._do_work()
 
        # Thread main loop, run until quit
        self._root.mainloop()

        # Thread cleanup. Ensure root is deleted here to avoid getting
        # Tcl_AsyncDelete: async handler deleted by the wrong thread
        self._root = None
        import gc
        gc.collect()

    def _shutdown(self):
        self._root.destroy()


_Tk_thread = None

def Tk_thread(fn=None, discard=False) -> futures.Executor:
    """
        All Tk GUI interactions must take place on the thread on which the root object 
        was created. This function returns a *concurrent.futures.Executor* which creates   
        the root object and can run commands on the dedicated Tk thread. It can also be 
        used a decorator.
        
        **All Tk interaction must take place via this thread**.

        Examples
        --------
            
        ::
            
            future = Tk_thread().submit(my_func, my_args)
            future.result() # blocks

            @Tk_thread(discard=True)
            def myfunc():
                # This is run on the Tk thread and returns a *future* to the result.
                # By specifying `discard=True` we tell xlOil that we're not going to
                # keep track of that future and so it should log any exceptions.
                ... 

    """

    global _Tk_thread

    if _Tk_thread is None:
        _Tk_thread = TkExecutor()
        # Send this blocking no-op to ensure Tk is created on our thread now
        _Tk_thread.submit(lambda: 0).result()

    return _Tk_thread if fn is None else _Tk_thread._wrap(fn, discard)

# Create thread on import - I'm not necessarily a fan of this blocking!
Tk_thread()


from xloil.gui import CustomTaskPane

class TkThreadTaskPane(CustomTaskPane):
    """
        Wraps a Tk window to create a CustomTaskPane object. 
    """

    def __init__(self):
        """
        Wraps a  to create a CustomTaskPane object. The ``draw_widget`` function
        is executed on the `xloil.gui.Tt_thread` and is expected to return a object.
        """
        super().__init__()

        self.contents = self.draw().result() # Blocks

        # Calling winfo_id() on a tk.Toplevel gives the hWnd of a child which 
        # represents the window client area. We want the *actual* top level window
        # which is the parent.
        from ctypes import windll
        self._hwnd = Tk_thread().submit(
            lambda: windll.user32.GetParent(self.contents.winfo_id())).result()

    def hwnd(self):
        return self._hwnd
             
    def on_destroy(self):
        Tk_thread().submit(lambda: self.contents.destroy())
        super().on_destroy()

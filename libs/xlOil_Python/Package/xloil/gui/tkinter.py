"""
    You must import this module before any other mention of `tkinter`: this allows xlOil to 
    create a thread to manage the Tk GUI and the Tk root object.  *All* interaction with the 
    *Tk* must be done on that thread or crashes will ensue.  Use `Tk_thread.submit(...)`
    or the `@Tk_thread` to ensure functions are run on the thread.
"""

import xloil
from xloil.gui import CustomTaskPane, _GuiExecutor, _ConstructInExecutor
from xloil._core import XLOIL_EMBEDDED

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
            
        self._make_ready()

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

if XLOIL_EMBEDDED:
    _Tk_thread = TkExecutor()
    # Create thread on import - I'm not necessarily a fan of this blocking!
    _Tk_thread.submit(lambda: 0).result()

def Tk_thread(fn=None, discard=False) -> TkExecutor:
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

    return _Tk_thread if fn is None else _Tk_thread._wrap(fn, discard)



# Safe now we've created the Tk_thread
import tkinter

class TkThreadTaskPane(CustomTaskPane, metaclass=_ConstructInExecutor, executor=Tk_thread):
    """
        Wraps a Tk window to create a CustomTaskPane object.
    """

    def __init__(self):
        self._top_level = tkinter.Toplevel()
        self._top_level.withdraw() # Create hidden

    @property
    def top_level(self) -> tkinter.Toplevel:
        """
            This returns a *tkinter.Toplevel* window into which the pane's contents
            should be placed.
        """
        return self._top_level

    def _get_hwnd(self):
        # Calling winfo_id() on a tk.Toplevel gives the hWnd of a child which 
        # represents the window client area. We want the *actual* top level window
        # which is the parent. Note we don't combine self.draw and getting
        # the hwnd in one call because that doesn't work for some reason...
        from ctypes import windll
        return Tk_thread().submit(
            lambda: (windll.user32.GetParent(self._top_level.winfo_id()), True))

    def on_destroy(self):
        super().on_destroy()
        Tk_thread().submit(lambda: self._top_level.destroy())


from .tk_console import TkConsole

import tkinter as tk
import time
import gc
import threading
import xloil

class _TkinterGUI:

    def __init__(self, start):
        if start:
            self._thread = threading.Thread(target=self._main_loop)
            self._thread.start()
            time.sleep(1)  # wait for self._root?
       
    def __del__(self):
        if self._root is not None:
            self._root.after(0, self._root.destroy)
            tk.wait_window(self._root)
            self._thread.join()
        
    def _main_loop(self):
        try:
            self._root = tk.Tk()
            self._root.withdraw()
            self._root.mainloop()
        
            # Avoid Tcl_AsyncDelete: async handler deleted by the wrong thread
            self._root = None
            gc.collect()
        except e:
            xloil.log(e.msg, level="error")

    @property
    def root(self):
        return self._root

# Not currently required
# tkinter_gui = _TkinterGUI()

_tkinter_root = None

def tk_root():
    global _tkinter_root
    # Only create Tk once - doing this multiple times will coredump tcl
    if _tkinter_root is None:
        import tkinter as tk
        _tkinter_root = tk.Tk(baseName="xlOil")
        _tkinter_root.withdraw()

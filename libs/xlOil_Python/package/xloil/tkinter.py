import queue
import xloil
import concurrent.futures as futures
import concurrent.futures.thread
import threading

# TODO: this looks an awful lot like a copy of QtExecutor!
class TkExecutor(futures.Executor):

    def __init__(self):
        self._work_queue = queue.SimpleQueue()
        self._thread = threading.Thread(target=self._main_loop, name="TkGuiThread")
        self._broken = False
        self._root = None
        self._thread.start()

    def submit(self, fn, *args, **kwargs):
        if self._broken:
            raise futures.BrokenExecutor(self._broken)

        f = futures.Future()
        w = concurrent.futures.thread._WorkItem(f, fn, args, kwargs)

        self._work_queue.put(w)
        return f

    def shutdown(self, wait=True, cancel_futures=False):
        if not self._broken:
            self.submit(self._root.destroy)
    
    @property
    def root(self):
        return self._root

    def _do_work(self):
        try:
            while True:
                work_item = self._work_queue.get_nowait()
                if work_item is not None:
                    work_item.run()
                    del work_item
        except queue.Empty:
            return
        finally:
            self._root.after(200, self._do_work)

    def _main_loop(self):

        try:
            import tkinter as tk

            self._root = tk.Tk(baseName="xlOil")
            self._root.withdraw()
            
             # Trigger timer to run any pending queue items now
            self._do_work()
 
            # Thread main loop, run until quit
            self._root.mainloop()

            # Thread cleanup
            # Avoid Tcl_AsyncDelete: async handler deleted by the wrong thread
            self._root = None
            import gc
            gc.collect()
            self._broken = True

        except Exception as e:
            self._broken = True
            xloil.log(f"TkThread failed: {e}", level='error')


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
        # PyBye is called before `threading` module teardown, whereas `atexit` comes later
        xloil.event.PyBye += _Tk_thread.shutdown
        # Send this blocking no-op to ensure Tk is created on our thread
        # before we proceed, otherwise Tk may try to create one elsewhere
        _Tk_thread.submit(lambda: 0).result()

    return _Tk_thread

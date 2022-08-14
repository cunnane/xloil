import xloil


class _Handler_debugpy:

    def __init__(self):    
        import debugpy
        debugpy.listen(5678)
    def call(self, type, value, trace):
        import debugpy
        # This doesn't actually do post-mortem debuggin so it fairly useless
        # It just breaks on a different python thread!
        debugpy.breakpoint()


class _Handler_pdb_window:

    def __init__(self):
        ...

    def call(self, type, value, trace):

        from xloil.gui.tkinter import Tk_thread

        console = Tk_thread().submit(self._open_console, Tk_thread().root, trace)
        console.result() # Blocks

    @staticmethod
    def _open_console(tk_root, trace):

        import tkinter as tk
        import sys
        
        #
        # Borrowed from
        # https://stackoverflow.com/questions/21811464/
        #
        class Console(tk.Frame):
            def __init__(self, parent, invoke):
                tk.Frame.__init__(self, parent)
                self.parent = parent

                self.real_std_in_out = (sys.stdin, sys.stdout, sys.stderr)

                sys.stdout = self
                sys.stderr = self
                sys.stdin = self

                import queue
                self.stdin_buffer = queue.Queue()

                self.create_widgets()

                from threading import Thread

                def run():
                    try:
                        invoke()
                    finally:
                        from .tkinter import Tk_thread
                        Tk_thread().submit(self.destroy)

                self._console_thread = Thread(target=run)
                self._console_thread.start()


            def destroy(self):
                self.exit = True
                self.stdin_buffer.put("\n\nexit()\n")
                sys.stdin, sys.stdout, sys.stderr = self.real_std_in_out
                super().destroy()

            def enter(self, event):
                input_line = self.ttyText.get("input_start", "end")
                self.ttyText.mark_set("input_start", "end-1c")
                self.ttyText.mark_gravity("input_start", "left")
                self.stdin_buffer.put(input_line)

            def write(self, string):
                self.ttyText.insert('end', string)
                self.ttyText.mark_set("input_start", "end-1c")
                self.ttyText.see('end')

            def create_widgets(self):
                self.ttyText = tk.Text(self.parent, wrap='word')
                self.ttyText.pack(expand=True, fill=tk.BOTH)
                self.ttyText.bind("<Return>", self.enter)
                self.ttyText.mark_set("input_start", "end-1c")
                self.ttyText.mark_gravity("input_start", "left")

            def flush(self):
                pass

            def readline(self):
                line = self.stdin_buffer.get()
                return line


        top_level = tk.Toplevel(tk_root)

        def disable_debugging():
            xloil.event.UserException.clear()

        menu = tk.Menu(top_level)
        menu.add_command(label="Disable Debugger", command=disable_debugging)
        top_level.config(menu=menu)

        import pdb
        main_window = Console(top_level, lambda: pdb.post_mortem(trace))
        main_window.deiconify()
        
        
_exception_handler = None

def _handler_func(type, value, trace):
    # Don't pop the debugger up whilst the user is trying to enter function
    # args in the wizard: that's just rude!
    if not xloil.in_wizard():
        _exception_handler.call(type, value, trace)

def exception_debug(debugger):
    """
        Selects a debugger for exceptions in user code. Only effects exceptions
        which are handled by xlOil. Choices are:

        **'pdb'**
        
        opens a console window with pdb active at the exception point

        **None**
        
        Turns off exception debugging

    """
    global _exception_handler

    if debugger is None:
        _exception_handler = None
        xloil.event.UserException.clear()
        return

    handlers = {
        'pdb': _Handler_pdb_window,
        'vs': _Handler_debugpy
    }

    if not debugger in handlers:
        raise Exception(f"Unsupported debugger {debugger}. Choose from: {handlers.keys()}")
    
    _exception_handler = handlers[debugger]()

    # No more than one exception handler, so clear event first
    xloil.event.UserException.clear()
    xloil.event.UserException += _handler_func

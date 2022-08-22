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
        
        from xloil.gui.tkinter import TkConsole

        top_level = tk.Toplevel(tk_root)

        def disable_debugging():
            xloil.event.UserException.clear()

        menu = tk.Menu(top_level)
        menu.add_command(label="Disable Debugger", command=disable_debugging)
        top_level.config(menu=menu)

        import pdb
        console = TkConsole(top_level, lambda: pdb.post_mortem(trace),
            fg='white', bg='black', font='Consolas', insertbackground='red')
        console.pack()
        console.bind("<<CommandDone>>", lambda e: top_level.destroy())

        top_level.deiconify()
        
        
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

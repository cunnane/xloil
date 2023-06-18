import xloil

class _Handler_pdb_window:

    def __init__(self):
        ...

    def call(self, type, value, trace):

        # Don't pop the debugger up whilst the user is trying to enter function
        # args in the wizard: that's just rude!
        if xloil.in_wizard():
            return

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
        
        # Auto resize widget with window
        console.pack(expand=True, fill=tk.BOTH)

        # Destroy window if pdb exits
        console.bind("<<CommandDone>>", lambda e: top_level.destroy())

        top_level.deiconify()
        

def debugpy_listen(port=None):
    
    if _is_debugpy_listening():
        return
 
    if port is None:
        addin = xloil.source_addin() 
        port = int(addin.settings['xlOil_Python']['DebugPyPort'])

    import debugpy
    connection = debugpy.listen(port)

    xloil.log(f"Debugpy listening on {connection}")

    def enable_debug():
        import debugpy, threading
        xloil.log.debug(f"Debugpy enabling debugging on thread {threading.get_native_id()}")
        #debugpy.debug_this_thread()
        debugpy.trace_this_thread()


    # Set main thread listener
    xloil.excel_callback(enable_debug)

    # This enables debugging of async/rtd functions running on the event loop.
    # Except it doesn't.  VS code refuses to hit breakpoints in them. No idea
    # why, guess it's just buggy.
    xloil.get_async_loop().call_soon_threadsafe(enable_debug)

    # Add listener for addin background workers
    for addin in xloil.xloil_addins.values():
        addin.event_loop.call_soon_threadsafe(enable_debug)
    
    # We monkey-patch all existing threaded functions!
    for addin in xloil.xloil_addins.values():
        registrations = addin.functions()
        for f in registrations:
            if f.is_threaded:
                f.func = _debugpy_activate_thread_decorator(f.func)


def _is_debugpy_listening() -> bool:
    import sys
    if not 'debugpy' in sys.modules:
        return False

    import debugpy
    return debugpy.server.api._adapter_process is not None

def _debugpy_activate_thread():

    if not _is_debugpy_listening():
        return

    import debugpy
    debugpy.debug_this_thread()

def _debugpy_activate_thread_decorator(fn):
    """
    If debugpy.listen has not been invoked, just returns its argument
    """

    if not _is_debugpy_listening():
        return fn

    import functools, debugpy

    already_called = False

    @functools.wraps(fn)
    def wrap(*args, **kwargs):
        nonlocal already_called
        if not already_called:
            debugpy.debug_this_thread()
            already_called = True
        return fn(*args, **kwargs)

    return wrap


DEBUGGERS = ['', 'pdb', 'debugpy', 'vscode']

# Keeps a reference to the debug handler used as events only hold weak references
_PDB_HANDLER = _Handler_pdb_window().call

def use_debugger(debugger, **kwargs):
    """
        Selects a debugger for exceptions in user code. Only effects exceptions
        which are handled by xlOil. Choices are:

        **pdb**
        
        opens a console window with pdb active at the exception point

        **vscode** or **debugpy**

        listens

        **None**
        
        Turns off exception debugging (does not turn off debugpy)

    """

    if debugger is None or debugger == '':
        xloil.event.UserException.clear()
        return

    if debugger == 'pdb':
        xloil.log.debug("Setting pdb as exception post-mortem debugger")

        # No more than one exception handler: clear event first
        xloil.event.UserException.clear()
        xloil.event.UserException += _PDB_HANDLER

    elif debugger == "vscode" or debugger == "debugpy":
        debugpy_listen(**kwargs)

    else:
        raise KeyError(f"Unsupported debugger {debugger}. Choose from: {DEBUGGERS}")

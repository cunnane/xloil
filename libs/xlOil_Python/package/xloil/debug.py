import xloil

class _Handler_ptvsd:

    def __init__(self):    
        import ptvsd
        ptvsd.enable_attach()
    def call(self, type, value, trace):
        import ptvsd
        # Probably don't want this as it pauses Excel whilst waiting!
        #ptvsd.wait_for_attach()
        ptvsd.break_into_debugger()

class _Handler_pdb_window:

    def __init__(self):
        import tkinter as tk
        pass

    def call(self, type, value, trace):
        import pdb
        import tkinter as tk
        import sys
        from threading import Thread
        import queue


        root = tk.Tk(baseName="xlOil")

        #
        # Borrowed from
        # https://stackoverflow.com/questions/21811464/how-can-i-embed-a-python-interpreter-frame-in-python-using-tkinter
        #
        class Console(tk.Frame):
            def __init__(self, parent, exit_callback, console_invoke):
                tk.Frame.__init__(self, parent)
                self.parent = parent
                self.exit_callback = exit_callback
                self.destroyed = False

                self.real_std_in_out = (sys.stdin, sys.stdout, sys.stderr)

                sys.stdout = self
                sys.stderr = self
                sys.stdin = self

                self.stdin_buffer = queue.Queue()

                self.createWidgets()

                self.consoleThread = Thread(target=lambda: self.run_console(console_invoke))
                self.consoleThread.start()

            def run_console(self, func):
                try:
                    func()
                except SystemExit:
                    if not self.destroyed:
                        self.after(0, self.exit_callback)

            def destroy(self):
                self.stdin_buffer.put("\n\nexit()\n")
                self.destroyed = True
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

            def createWidgets(self):
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

        main_window = Console(root, root.destroy, lambda: pdb.post_mortem(trace))
        main_window.mainloop()

_exception_handler = None

def _handler_func(type, value, trace):
    _exception_handler.call(type, value, trace)

def exception_debug(debugger):
    """
    Selects a debugger for exceptions in user code. Only effects exceptions
    which are handled by xlOil. Choices are:

        pdb 
        ---
        opens a console window with pdb active at the exception point

        vs 
        --
        uses ptvsd (Python Tools for Visual Studio) to enable Visual Studio
        (or VS Code) to connect via a remote session. Connection is on the default 
        settings i.e. localhost:5678. This means your lauch.json in VS Code should be:

        ::

            {
                "name": "Attach (Local)",
                "type": "python",
                "request": "attach",
                "localRoot": "${workspaceRoot}",
                "port": 5678,
                "host":"localhost"
            }

        A breakpoint is also set a the exception site.

        None
        ----
        Turns off exception debugging

    """
    global _exception_handler

    if debugger is None:
        _exception_handler = None
        xloil.event.UserException.clear()
        return

    handlers = {
        'pdb': _Handler_pdb_window,
        'vs': _Handler_ptvsd
    }

    if not debugger in handlers:
        raise Exception(f"Unsupported debugger {debugger}. Choose from: {handlers.keys()}")
    
    _exception_handler = handlers[debugger]()

    # No more than one exception handler, so clear event first
    xloil.event.UserException.clear()
    xloil.event.UserException += _handler_func
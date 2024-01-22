import sys
import tkinter

# TODO: use sys.__stdin__
_REAL_STD_IN_OUT = (sys.stdin, sys.stdout, sys.stderr)

#
# Created with help from https://stackoverflow.com/questions/21811464/
#

class TkConsole(tkinter.Frame):
    """
        A *very* simple console-in-a-text-box. It 'steals' stdin/out/err to 
        allow python REPLs like pdb.post_mortem or code.interact to work
        in-process. Because of this stealing, opening more than one Console
        will lead to unexpected results.

        It has no terminal emulation, but it does support command history. 

        The widget emits an event <<CommandDone>> when the provided invoke
        command completes (typically because the REPL is quit by the user)

        Parameters
        ----------
        
        parent:
            Parent widget
        invoke:
            The command to run. If this doesn't start a REPL, then I'm not
            sure you should be using this class!
        kwargs:
            All other kwargs are passed to the tkinter.Text constructor.
    """

    MARK_INPUT = "input_start"

    def __init__(self, parent, invoke, **kwargs):
        super().__init__(parent)
  
        self._destroyed = False

        import queue
        self.stdin_buffer = queue.Queue()
        
        # History of commands entered
        self._history = []
        self._history_pos = 0

        self._text = self._create_text_widget(kwargs)
        
        def run():
            try:
                import sys
                previous_std_in_out = (sys.stdin, sys.stdout, sys.stderr)
                sys.stdout = self
                sys.stderr = self
                sys.stdin  = self
                invoke()
            finally:
                # We avoid resetting the std streams if another instance has already 
                # restored them to the system defaults. This may happen if more than 
                # one Console was opened (inadvisable), then closed in a different order.
                if not sys.stdin is _REAL_STD_IN_OUT[0]:
                    sys.stdin, sys.stdout, sys.stderr = previous_std_in_out
                if not self._destroyed:
                    self.event_generate("<<CommandDone>>")

        from threading import Thread
        self._console_thread = Thread(target=run)
        self._console_thread.start()

    def destroy(self):
        self._destroyed = True
        self.stdin_buffer.put("\n\nexit()\n")
        super().destroy()

    def _press_enter(self, event):
        
        input_line = self._text.get(self.MARK_INPUT, "end")

        # Add a newline as we stop the normal event handler from being called
        self._text.insert('end', '\n')
        
        # Send input to our stdin
        self.stdin_buffer.put(input_line)
        
        # Add to history unless it's a blank line
        stripped_input = input_line.strip()
        if len(stripped_input) > 0:
            self._history.append(stripped_input)
            self._history_pos = len(self._history)
        
        return 'break'

    def _press_updown(self, step:int):
        # Increment history pointer
        self._history_pos = (self._history_pos - step) % len(self._history)
        text = self._text
        # Remove existing input
        text.delete(self.MARK_INPUT, 'end')
        # Copy history entry to input mark
        text.insert(self.MARK_INPUT, self._history[self._history_pos])
        # Move caret to start of input
        text.mark_set("insert", self.MARK_INPUT)
        # Do not let default handler move caret
        return 'break'

    def _press_left(self, event):
        # Block attempts to move the caret before the input start mark
        text = self._text
        if text.index("insert") == text.index(self.MARK_INPUT):
            return 'break'

    def _press_up(self, event):
        return self._press_updown(1)

    def _press_down(self, event):
        return self._press_updown(-1)

    def _create_text_widget(self, options):
        """
        Creates a (scrolled) text widget with the following key overrides:
              * Up/Down arrows insert command history
              * Backspace / left arrow cannot move beyond input prompt
              * Enter sends command to the console thread  
        """
        from tkinter.scrolledtext import ScrolledText
        text = ScrolledText(self, wrap='word', **options)
        text.pack(expand=True, fill=tkinter.BOTH)

        text.bind("<Return>", self._press_enter)
        text.bind("<Up>", self._press_up)
        text.bind("<Down>", self._press_down)
        text.bind("<Left>", self._press_left)
        text.bind("<BackSpace>", self._press_left)

        # Set input mark to 1 char before end (to avoid newline char)
        text.mark_set(self.MARK_INPUT, "end-1c")
        # Gravity 'left' means the mark maintains position as new text is inserted at 
        # its location
        text.mark_gravity(self.MARK_INPUT, "left")

        return text

    def write(self, string):
        text = self._text
        text.insert('end', string)
        # Move input mark to 1 char before end (to avoid newline char)
        text.mark_set(self.MARK_INPUT, "end-1c")
        # Move caret to same location
        text.mark_set("insert", "end-1c")
        # Scroll to end 
        text.see('end')

    def flush(self):
        pass

    def readline(self):
        line = self.stdin_buffer.get()
        return line


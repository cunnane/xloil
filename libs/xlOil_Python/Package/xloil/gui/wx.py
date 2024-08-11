"""
    You must import this module before any other mention of `wx`: this allows xlOil to 
    create a thread to manage the wx GUI and the wx root object.  *All* interaction with the 
    *wx* must be done on that thread or crashes will ensue.  Use `wx_thread.submit(...)`
    or the `@wx_thread` to ensure functions are run on the thread.
"""

import xloil
from xloil.gui import CustomTaskPane, _GuiExecutor, _ConstructInExecutor
from xloil._core import XLOIL_EMBEDDED
from xloil import log


class WxExecutor(_GuiExecutor):

    def __init__(self):
        self._app = None
        super().__init__("WxGuiThread")

    @property
    def app(self):
        return self._app

    def _do_work(self):
        super()._do_work()
        import wx
        wx.CallLater(200, self._do_work)

    def _main(self):

        import wx

        self._app = wx.App()
        
        # Mainloop will immediately exit without at least 1 window
        hidden = wx.Frame(None, title="Hello World")

        self._make_ready()

        # Run any pending queue items now
        self._do_work()
 
        # Thread main loop, run until quit
        self._app.MainLoop()
        
        log.debug("Wx: Finished running main loop, finalising Application")

        # See https://stackoverflow.com/questions/49304429/
        wx.DisableAsserts()

        for window in wx.GetTopLevelWindows():
            window.Destroy()

        self._app = None


    def _shutdown(self):
        self._app.ExitMainLoop()


_wx_thread = None

if XLOIL_EMBEDDED:
    _wx_thread = WxExecutor()
    # Send this blocking no-op to ensure wx is created on our thread now
    _wx_thread.submit(lambda: 0).result()

def wx_thread(fn=None, discard=False) -> WxExecutor:
    """
        All wx GUI interactions must take place on the thread on which the root object 
        was created. This function returns a *concurrent.futures.Executor* which creates   
        the root object and can run commands on the dedicated wx thread. It can also be 
        used a decorator.
        
        **All wx interaction must take place via this thread except CallLater**.

        Examples
        --------
            
        ::
            
            future = wx_thread().submit(my_func, my_args)
            future.result() # blocks

            @wx_thread(discard=True)
            def myfunc():
                # This is run on the wx thread and returns a *future* to the result.
                # By specifying `discard=True` we tell xlOil that we're not going to
                # keep track of that future and so it should log any exceptions.
                ... 

    """
    global _wx_thread
    return _wx_thread if fn is None else _wx_thread._wrap(fn, discard)

try:
    import wx
except ImportError:
    from ._core import XLOIL_READTHEDOCS
    if not XLOIL_EMBEDDED:
        class wx:
            class Frame:
                # Placeholder for wx.Frame
                ...

class WxThreadTaskPane(CustomTaskPane, metaclass=_ConstructInExecutor, executor=wx_thread):
    """
        Wraps a wx window to create a CustomTaskPane object.
    """

    def __init__(self, frame=None):
        
        if isinstance(frame, wx.Frame):
            self._frame = frame
        elif frame is not None:
            self._frame = frame()
        else:
            self._frame = wx.Frame()

    @property
    def frame(self) -> wx.Frame:
        """
            This returns a *wx.Frame* window into which the pane's contents
            should be placed.
        """
        return self._frame

    def _get_hwnd(self):
        def prepare(frame):
            frame.Show()
            # import wx
            # frame.Move(wx.Point(0,0))
            # Unfortunately, Wx windows can't be reparented, so return False
            return frame.GetHandle(), False
        return wx_thread().submit(prepare, self._frame)

    def on_destroy(self):
        super().on_destroy()
        wx_thread().submit(lambda: self._frame.Destroy())
        

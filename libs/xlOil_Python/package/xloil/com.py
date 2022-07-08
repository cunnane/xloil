from . import _core

class EventsPaused():
    """
    A context manager which stops Excel events from firing whilst
    the context is in scope
    """
    def __enter__(self):
        _core.event.pause()
        return self
    def __exit__(self, type, value, traceback):
        _core.event.allow()

    
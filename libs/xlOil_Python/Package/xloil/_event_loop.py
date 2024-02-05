from .logging import log_except

def _logged_wrapper(func):
    """
    Wraps func so that any errors are logged. Invoked from the core.
    """
    def logged_func(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            log_except(f"Error during {func.__name__}")
    return logged_func

async def _logged_wrapper_async(coro):
    """
    Wraps coroutine so that any errors are logged. Invoked from the core.
    """
    try:
        return await coro
    except Exception as e:
        log_except(f"Error during coroutine")

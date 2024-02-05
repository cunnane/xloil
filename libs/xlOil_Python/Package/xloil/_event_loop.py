
def _pump_message_loop(loop, timeout:float):
    """
    Called internally to run the asyncio message loop. Returns the number of active tasks
    """
    import asyncio

    async def wait():
        await asyncio.sleep(timeout)
    
    loop.run_until_complete(wait())

    all_tasks = asyncio.all_tasks if sys.version_info[:2] > (3, 6) else asyncio.Task.all_tasks
    return len([task for task in all_tasks(loop) if not task.done()])

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

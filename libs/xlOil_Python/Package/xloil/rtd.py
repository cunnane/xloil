import xloil as xlo

class RtdSimplePublisher(xlo.RtdPublisher):
    """
    Implements the boiler-plate involved in an RtdPublisher for the vanilla case.
    Just need to implement an exception-safe async run_task which collects the data
    and publishes it to the RtdServer.
    """

    def __init__(self, topic):
        # Must call this ctor explicitly or pybind will crash
        super().__init__()  
        self._topic = topic
        self._task = None
    
    async def run_task(topic):
        ...

    def connect(self, num_subscribers):
        if self.done():
            self._task = xlo.get_event_loop().create_task(run_task(self._topic))
                
    def disconnect(self, num_subscribers):
        if num_subscribers == 0:
            self.stop()
            return True  # Schedules the publisher for destruction
                
    def stop(self):
        if self._task is not None: 
            self._task.cancel()
        
    def done(self):
        return self._task is None or self._task.done()
            
    def topic(self):
        return self._topic


def subscribe(server:xlo.RtdServer, topic:str, coro):
    """
    Subscribes to `topic` on the given RtdServer, starting the publishing
    task `coro` if no publisher for `topic` is currently running.

    `coro` must be a coroutine and the string `topic` must be derived from
    the function's arguments in a way which uniquely identifies the data to 
    be published 
    """
    if server.peek(topic) is None:
        from .register import async_wrapper, find_return_converter
        from .func_inspect import Arg

        func_args, return_type = Arg.full_argspec(coro)
        if any(func_args):
            raise ValueError("Cororoutine should have zero args")

        return_converter = find_return_converter(return_type)

        wrapped = async_wrapper(coro)

        server.start_task(topic, wrapped, return_converter)

    return server.subscribe(topic)  
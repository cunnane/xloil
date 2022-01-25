import xloil as xlo


class RtdSimplePublisher(xlo.RtdPublisher):
    """
    Implements the boiler-plate involved in an RtdPublisher for the vanilla case.
    Just need to implement an exception-safe async run_task which collects the data
    and publishes it to the RtdServer.
    """

    def __init__(self, topic):
        # Must call this ctor explicitly or the pybind will crash
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
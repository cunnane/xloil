import xloil as xlo
import IPython
import pickle
from IPython.display import publish_display_data
import jupyter_client
import os
import asyncio

def _serialise(obj):
    return pickle.dumps(obj, protocol=0).decode('utf-8')

def _deserialise(dump:str):
    return pickle.loads(dump.encode('utf-8'))

class MonitoredVariables:

    def __init__(self, ipy_shell):
        self._values = dict()
        self._shell = ipy_shell

        ipy_shell.events.register('post_execute', self.post_execute)
    
    def post_execute(self):
        updates = {}
        for name, val in self._values.items():
            that_val = self._shell.user_ns.get(name, None)
            # Use is rather than == as the latter may not return a single value e.g. numpy
            if not that_val is val:
                updates[name] = that_val
                self._values[name] = that_val

        if len(updates) > 0:
            publish_display_data(
                { "xloil/data": _serialise(updates) },
                { 'type': "VariableChange" }
            )

    def watch(self, name):
        self._values[name] = globals().get(name, None)
        self.post_execute()

    def stop_watch(self, name):
        del self._values[name]

    def close(self):
        self._shell.events.unregister('post_execute', self.post_execute)


def _func_decorator(fn, *args, **kwargs):
    
    # TODO: Work out how to deal with the args later
    spec = xlo.FuncDescription(fn)

    # Hack the function description, replacing the function object with
    # its name
    spec._func = spec._func.__name__

    publish_display_data(
        { "xloil/data": _serialise(spec) },
        { 'type': "FuncRegister" }
    )

    return fn


def function_invoke(func, args_data, kwargs_data):
    args = _deserialise(args_data)
    kwargs = _deserialise(kwargs_data)
    result = func(*args, **kwargs)
    publish_display_data(
        { "xloil/data": _serialise(result) },
        { 'type': "FuncResult" }
    )
    return result # Not used, just in case tho

class _VariableWatcher(xlo.RtdPublisher):

    def __init__(self, var_name, topic_name, connection): 
        super().__init__()
        self._name = var_name
        self._topic = topic_name
        self._connection = connection ## TODO: weak-ptr?

    def connect(self, num_subscribers):
        if num_subscribers == 1:
            self._connection.execute(f"_xloil_vars.watch('{self._name}')", silent=True)

    def disconnect(self, num_subscribers):
        if num_subscribers == 0:
            self._connection.execute(f"_xloil_vars.stop_watch('{self._name}')", silent=True)
            return True # Schedule topic for destruction

    def stop(self):
        self._connection.stop_watch_variable(self._name)

    def done(self):
        return True

    def topic(self):
        return self._topic


class JupyterNotReadyError(Exception):
    pass

_rtdServer = xlo.RtdServer()

class _JupyterConnection:
    
    _pending_messages = dict() # str -> Future
    _watched_variables = dict()
    _registered_funcs = set()
    _ready = False # Indicates the connection is ready to receive commands  

    def __init__(self, connection_file, xloil_path):

        from jupyter_client.asynchronous import AsyncKernelClient

        self._loop = xlo.get_event_loop()
        self._client = AsyncKernelClient()
        self._xloil_path = xloil_path.replace('\\', '/')
        self._connection_file = connection_file

        cf = jupyter_client.find_connection_file(self._connection_file)
        xlo.log(f"Jupyter: found connection for file {self._connection_file}", level='trace')
        self._client.load_connection_file(cf)

    async def connect(self):
        
        self._client.wait_for_ready(timeout=3)
        # TODO: what if we timeout?

        # Flush shell channel (wait_for_ready doesn't do this for some reason), but we need it 
        # receive the shell message below
        while True:
            from queue import Empty
            try:
                msg = await self._client.get_shell_msg(timeout=0.2)
            except Empty:
                break
        

        # Setup the python path to find xloil_jupyter. We do a reload of xloil because
        # we override xloil.func - this allows repeated connection to the same kernel
        # without error.  We also connect the variable monitor

        # TODO: if the target notebook already has imported xloil under a different name
        # I suspect the overwrite of xloil.func will not work.
        msg_id = self._client.execute(
            "import sys, importlib, IPython\n" + 
            f"sys.path.append('{self._xloil_path}')\n" + 
            "import xloil\n"
            "import xloil_jupyter\n"
            "importlib.reload(xloil)\n"
            "xloil.func = xloil_jupyter._func_decorator\n" # Overide xloil.func to do our own thing
            "_xloil_vars = xloil_jupyter.MonitoredVariables(get_ipython())\n"
        )

        # TODO: retry?
        # Some examples online suggest get_shell_msg(msg_id) should work. It doesn't, which is 
        # a shame as it would be rather nice.
        msg = {}
        while msg.get('parent_header', {}).get('msg_id', None) != msg_id:
            msg = await self._client.get_shell_msg()
            if msg['content']['status'] == 'error':
                raise Exception(f"Connection failed: {msg['content']['evalue']}")

        self._sessionId = msg['header']['session']
        self._watched_variables.clear()
        self._ready = True


    def close(self):

        # If still ready, i.e. not triggered by a kernel restart, clean up our variables
        if self._ready:
            self._client.execute("_xloil_vars.close()\ndel _xloil_vars")

        self._ready = False

        # Copy topics because the disconnect for a watched variable will
        # remove it from the dict
        variable_topics = [x.topic() for x in self._watched_variables.values()]
        
        for topic in variable_topics:
            _rtdServer.drop(topic)
        
        # Remove any stragglers
        self._watched_variables.clear()

        for func_name in self._registered_funcs:
            xlo.deregister_functions(None, func_name)

    async def wait_for_restart(self):

        session = self._sessionId

        while True:
            # We'd expect to see exec_state == 'starting', on iopub but sometimes it doesn't 
            # appear (???) so we settle for any exec_state with a new session id.
            msg = await self._client.get_iopub_msg()
            #xlo.log(f"WFR: {msg}")
            exec_state = msg.get('content', {}).get('execution_state', None)
            session = msg['header']['session']
            if exec_state is not None and session != self._sessionId:
                break
            await asyncio.sleep(1)

        xlo.log(f"Restart: new session {session}", level='trace')

        await self.connect()

    def execute(self, *args, **kwargs):
        if not self._ready:
            raise JupyterNotReadyError()
        return self._client.execute(*args, **kwargs)

    async def invoke(self, func_name, *args, **kwargs):
        if not self._ready:
            raise JupyterNotReadyError()

        args_data = repr(_serialise(args))
        kwargs_data = repr(_serialise(kwargs))
        msg_id = self.execute("xloil_jupyter.function_invoke("
           f"{func_name}, {args_data}, {kwargs_data})")
        future = self._loop.create_future()
        self._pending_messages[msg_id] = future
        return await future

    def _watch_prefix(self, name):
        prefix = self._client.get_connection_info()["key"].decode('utf-8')
        return f"{prefix}_{name}" 

    def watch_variable(self, name):
        if not self._ready:
            raise JupyterNotReadyError()

        topic = self._watch_prefix(name)
        if not name in self._watched_variables:

            # TODO: I think that if we don't retain a ref to the watcher it will somehow get
            # lost by pybind, but can we check this again?
            watcher = _VariableWatcher(name, topic, self)
            self._watched_variables[name] = watcher

            xlo.log(f"Starting variable watch {name}", level='trace')
            _rtdServer.start(watcher)

        return _rtdServer.subscribe(topic)

    def stop_watch_variable(self, name):
        try:
            del self._watched_variables[name]
        except KeyError:
            pass

    def publish_variables(self, updates:dict):
        for name, value in updates.items():
            _rtdServer.publish(self._watch_prefix(name), value)

    async def process_messages(self):
   
        while self._client.is_alive():
            from queue import Empty
            try:
                # At the moment we communicate over the public iopub channel 
                # TODO: consider using the private comm channel rather
                msg = await self._client.get_iopub_msg()
                content = msg['content']
            except (Empty, KeyError):
                # Timed out waiting for messages, or msg had no content
                continue 
            
            xlo.log(f"Jupyter Msg: {content}", level='trace')

            # Kernel is shutting down, break out of loop
            if msg['header']['msg_type'] == 'shutdown_reply':
                xlo.log(f"Jupyter kernel shutdown: {self._connection_file}", level='info')
                self._ready = False
                return content['restart']

            # Check if we have have an expected reply
            parent_id = msg.get('parent_header', {}).get('msg_id', None)
            pending = self._pending_messages.get(parent_id, None)

            # If we match a pending message and have an error, publish it. Non-errors 
            # are handled later
            if pending is not None and msg['header']['msg_type'] == 'error':
                err_msg = content['evalue']
                pending.set_result(err_msg)
            

            # Look for xlOil specific message content
            data = content.get('data', {})
            xloil_data = data.get('xloil/data', None)
            if xloil_data is None: 
                continue
            payload = _deserialise(xloil_data)
            meta_type = content['metadata']['type']
            
            xlo.log(f"Jupyter xlOil Msg: {payload}", level='trace')

            # Handle xlOil messages
            if meta_type == "VariableChange":
                self.publish_variables(payload)

            elif meta_type == "FuncRegister":
                descr = payload

                # The func description has be "hacked" so that _func is 
                # the string name of the function to be invoked rather
                # that a function object.  We use the name to create a
                # shim which invokes jupyter. This is set as the 
                # function object
                func_name = descr._func

                @xlo.async_wrapper
                async def shim(*args, **kwargs):
                    return await self.invoke(func_name, *args, **kwargs)

                # Correctly set the function description: it's also RTD 
                # async now
                descr._func = shim
                descr.rtd = True
                descr.is_async = True

                xlo.register_functions(None, [descr.create_holder()])

                # Keep track of our funtions for a clean shutdown
                self._registered_funcs.add(func_name)

            elif meta_type == "FuncResult":
                if pending is None:
                    xlo.log(f"Unexpected function result: {msg}")
                else:
                    pending.set_result(payload)
                    continue

            else:
                raise Exception(f"Unknown xlOil message: {meta_type}, {payload}")



    # TODO: support 'with'?
    #def __enter__(self):
    #def __exit__(self, exc_type, exc_value, traceback)


class _JupterTopic(xlo.RtdPublisher):

    def __init__(self, topic, connection_file, xloil_path): 
        super().__init__()
        self._connection = _JupyterConnection(connection_file, xloil_path)
        self._topic = topic
        self._task = None 
        self._cacheRef = None

    def connect(self, num_subscribers):

        if self.done():
            conn = self._connection

            async def run():
                try:
                    await conn.connect()
                    while True:
                        self._cacheRef = xlo.cache.add(conn, tag=self._topic)
                        # TODO: use a customer converter for this publish to stop xloil 
                        # unpacking the cacheref
                        _rtdServer.publish(self._topic, self._cacheRef)
                        restart = await conn.process_messages()
                        conn.close()
                        if not restart:
                            break
                        await conn.wait_for_restart()

                    _rtdServer.publish(self._topic, "KernelShutdown")

                except Exception as e:
                    _rtdServer.publish(self._topic, str(e))

            self._task = conn._loop.create_task(run())

    def disconnect(self, num_subscribers):
        if num_subscribers == 0:
            self.stop()

            # Cleanup our cache reference
            xlo.cache.remove(self._cacheRef)

            # Schedule topic for destruction
            return True 

    def stop(self):
        if not self._task is None:
            self._task.cancel()
        self._connection.close()

    def done(self):
        return self._task is None or self._task.done()

    def topic(self):
        return self._topic


@xlo.func(
    help="Connects to a jupyter (ipython) kernel given a connection file. Functions created"
         "in the kernel and decorated with xlo.func will be registered with Excel.",
    args={
            'ConnectionFile': "A file of the form 'kernel-XXX.json'. Generate it by executing the "
                              "'%connect_info' magic in a jupyter cell"
         }
)
def xloJpyConnect(ConnectionFile):
    topic = ConnectionFile.lower()
    if _rtdServer.peek(topic) is None:
        conn = _JupterTopic(topic, ConnectionFile, os.path.dirname(__file__))
        _rtdServer.start(conn)
    return _rtdServer.subscribe(topic)

@xlo.func(
    help="Fetches the value of the specifed variable from the given jupyter"
         "connection. Updates it live using RTD.",
    args={
        'Name': 'Case-sensitive name of a global variable in the jupyter kernel',
        'Connection': 'Connection ref output from xloJpyConnect'
    }
)
def xloJpyWatch(Name, Connection):
    return Connection.watch_variable(Name)

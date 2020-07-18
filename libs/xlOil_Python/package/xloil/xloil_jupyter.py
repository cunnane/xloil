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

#TODO: these messages seem like nice design, but they make the pickle bigger so kill them
class _VariableChangeMessage:
    def __init__(self, updates):
        self.updates = updates

class _FuncRegisterMessage:
    def __init__(self, description):
        self.description = description

class _FuncResultMessage:
    def __init__(self, result):
        self.result = result


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
                { "xloil/data": _serialise(_VariableChangeMessage(updates)) },
                { 'type': "VariableChange" }
            )

    def watch(self, name):
        self._values[name] = globals().get(name, None)
        self.post_execute()

    def stop_watch(self, name):
        del(self._values[name])

    def close(self):
        self._shell.events.unregister('post_execute', self.post_execute)


def _func_decorator(fn, *args, **kwargs):
    
    # TODO: Work out how to deal with the args later
    spec = xlo.FuncDescription(fn)

    # Hack the function description, replacing the function object with
    # its name
    spec._func = spec._func.__name__

    publish_display_data(
        { "xloil/data": _serialise(_FuncRegisterMessage(spec)) },
        { 'type': "FuncResult" }
    )

    return fn


def function_invoke(func, args_data, kwargs_data):
    args = _deserialise(args_data)
    kwargs = _deserialise(kwargs_data)
    result = func(*args, **kwargs)
    publish_display_data(
        { "xloil/data": _serialise(_FuncResultMessage(result)) },
        { 'type': "FuncResult" }
    )
    return result # Not used, just in case tho

class _VariableWatcher(xlo.RtdTopic):

    def __init__(self, var_name, topic_name, connection): 
        super().__init__()
        self._name = var_name
        self._topic = topic_name
        self._connection = connection ## TODO: weak-ptr?

    def connect(self, num_subscribers):
        if num_subscribers == 1:
            self._connection._client.execute(f"_xloil_vars.watch('{self._name}')", silent=True)

    def disconnect(self, num_subscribers):
        if num_subscribers == 0:
            self._connection._client.execute(f"_xloil_vars.stop_watch('{self._name}')", silent=True)
            return True # Schedule topic for destruction

    def stop(self):
        self._connection.stop_watch_variable(self._name)

    def done(self):
        return True

    def topic(self):
        return self._topic


_rtdManager = xlo.RtdManager()

class _JupyterConnection:
    
    _pending_messages = dict() # str -> Future
    _watched_variables = dict()
    _registered_funcs = set()

    def __init__(self, connection_file, xloil_path):

        from jupyter_client.asynchronous import AsyncKernelClient

        self._loop = xlo.get_event_loop()
        self._client = AsyncKernelClient()
        self._xloil_path = xloil_path.replace('\\', '/')
        self._connection_file = connection_file

        cf = jupyter_client.find_connection_file(connection_file)
        xlo.log(f"Jupyter: found connection for file {connection_file}")
        self._client.load_connection_file(cf)

        if not self._client.is_alive():
            raise Exception("Specified client is dead")

    async def connect(self):

        # Setup the python path to find xloil_jupyter
        # (Currently doing a reload to aid debugging)
        # Connect the variable monitor
        self._client.execute(
            "import sys, importlib, IPython\n" + 
            f"sys.path.append('{self._xloil_path}')\n" + 
            "import xloil\n"
            "import xloil_jupyter\n"
            "importlib.reload(xloil)\n" # Because we override xloil.func, we need to reset that
            "xloil.func = xloil_jupyter._func_decorator\n" # Overide xloil.func to do our own thing
            "_xloil_vars = xloil_jupyter.MonitoredVariables(get_ipython())\n"
        )

        msg = await self._client.get_shell_msg()
        self._sessionId = msg['header']['session']
        if 'content' in msg:
            content = msg['content']
            if content['status'] == 'error':
                raise Exception(f"Connection failed: {content['evalue']}")

    def close(self):
        self._client.execute("_xloil_vars.close()\ndel _xloil_vars")

        # Copy because the disconnect for a watched variable will
        # remove it from this dict
        variable_topics = [x.topic() for x in self._watched_variables.values()]
        
        for topic in variable_topics:
            _rtdManager.drop(topic)

        for func_name in self._registered_funcs:
            xlo.deregister_functions(None, func_name)

    async def invoke(self, func_name, *args, **kwargs):
        args_data = repr(_serialise(args))
        kwargs_data = repr(_serialise(kwargs))
        msg_id = self._client.execute("xloil_jupyter.function_invoke("
           f"{func_name}, {args_data}, {kwargs_data})")
        future = self._loop.create_future()
        self._pending_messages[msg_id] = future
        return await future

    def _watch_prefix(self, name):
        prefix = self._client.get_connection_info()["key"].decode('utf-8')
        return f"{prefix}_{name}" 

    def watch_variable(self, name):
        topic = self._watch_prefix(name)
        if name not in self._watched_variables:
            watcher = _VariableWatcher(name, topic, self)
            self._watched_variables[name] = watcher
            _rtdManager.start(watcher)
        return _rtdManager.subscribe(topic)

    def stop_watch_variable(self, name):
        del self._watched_variables[name]

    def publish_variables(self, updates:dict):
        for name, value in updates.items():
            _rtdManager.publish(self._watch_prefix(name), value)

    async def process_messages(self):
   
        while self._client.is_alive():
            from queue import Empty
            try:
                # At the moment we communicate over the public iopub channel 
                # TODO: consider using the private comm channel rather
                msg = await self._client.get_iopub_msg()
                content = msg.get('content', None)
                if content is None:
                    continue
                
                #xlo.log(f"Msg: {msg}", level='trace')

                parent_id = msg.get('parent_header', {}).get('msg_id', None)
                pending = self._pending_messages.get(parent_id, None)

                if pending is not None and msg['header']['msg_type'] == 'error':
                    err_msg = content['evalue']
                    pending.set_result(err_msg)
                
                data = content.get('data', {})
                xloil_data = data.get('xloil/data', None)
                if xloil_data is None: continue
     
                #meta_type = content['metadata']['type']
                xloil_msg = _deserialise(xloil_data)
                
                xlo.log(f"Jupyter xlOil Msg: {xloil_msg}", level='trace')

                if type(xloil_msg) is _VariableChangeMessage:
                    self.publish_variables(xloil_msg.updates)

                elif type(xloil_msg) is _FuncRegisterMessage:
                    descr = xloil_msg.description

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

                elif type(xloil_msg) is _FuncResultMessage:
                    if pending is None:
                        xlo.log(f"Unexpected function result: {msg}")
                    else:
                        pending.set_result(xloil_msg.result)
                        continue

                else:
                    raise Exception(f"Unknown message: {xloil_msg}")

            except Empty:
                pass

    # TODO: support 'with'?
    #def __enter__(self):
    #def __exit__(self, exc_type, exc_value, traceback)


class _JupterTopic(xlo.RtdTopic):

    def __init__(self, topic, connection_file, xloil_path): 
        super().__init__()
        self._connection = _JupyterConnection(connection_file, xloil_path)
        self._topic = topic
        self._task = None 
        self._cacheRef = None

    def connect(self, num_subscribers):
        if num_subscribers == 1:
            conn = self._connection
            async def run():
                await conn.connect()
                self._cacheRef = xlo.cache.add(conn, tag=self._topic)
                _rtdManager.publish(self._topic, self._cacheRef)
                await conn.process_messages()

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
    help="Connects to a jupyter kernel using the kernel-XXX.json file "
         "generated by executing the %connect_info magic")
def xloJpyConnect(ConnectionFile):
    topic = ConnectionFile.lower()
    if _rtdManager.peek(topic) is None:
        conn = _JupterTopic(topic, ConnectionFile, os.path.dirname(__file__))
        _rtdManager.start(conn)
    return _rtdManager.subscribe(topic)

@xlo.func(
    help="Fetches the value of the specifed variable from the given jupyter"
         "connection. Updates live using RTD.")
def xloJpyWatch(Name, Connection):
    return Connection.watch_variable(Name)

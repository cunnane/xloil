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

class _VariableChangeMessage:
    def __init__(self, updates):
        self.updates = updates

class MonitoredVariables:

    def __init__(self, ipy_shell):
        self._values = dict()
        self._shell = ipy_shell

        ipy_shell.events.register('post_execute', self.post_execute)
    
    def post_execute(self):
        self._check(self._shell.user_ns)

    def _check(self, source):
        updates = {}
        for name, val in self._values.items():
            that_val = source.get(name, None)
            if that_val != val:
                updates[name] = that_val
                self._values[name] = that_val

        if len(updates) > 0:
            publish_display_data(
                {
                    "xloil/data": _serialise(_VariableChangeMessage(updates))
                },
                {
                    'type': "VariableChange"
                }
            )

    def watch(self, name, value):
        self._values[name] = value

    def stop_watch(self, name):
        del(self._values[name])

class _FuncRegisterMessage:
    def __init__(self, description):
        self.description = description

def _ipython_decorator(f):
    spec = xlo._FuncDescription(f)
    #
    # Question: how is this going to work with xloil.arg?
    #

    # "Hack" the function description 
    spec._func = spec._func.__name__

    dump = _serialise(spec)
    publish_display_data(
        {'xloil/data': dump},
        {'type': "TestFuncSpec"}
    )

def _function_invoke(func, args_data, kwargs_data):
    args = _deserialise(args_data)
    kwargs = _deserialise(kwargs_data)
    result = func(*args, **kwargs)
    return _serialise(result)


#def _function_dispatch(connection, func_name, *args, **kwargs):
#    loop = asyncio.get_event_loop()
#    result = await connection._invoke(func_name, args, kwargs)
#    return result
        
class _VariableWatcher(xlo.RtdTopic):

    def __init__(self, var_name, kernel_client):
        super().__init__()
        self._name = var_name
        self._client = kernel_client

    def connect(self, num_subscribers):
        if num_subscribers == 1:
            self._client.execute(f"_xloil_vars.watch('{self._name}', {self._name})", silent=True)

    def disconnect(self, num_subscribers):
        if num_subscribers == 0:
            self._client.execute(f"_xloil_vars.stop_watch('{self._name}')", silent=True)

    def stop(self):
        pass # TODO: remove ourselves from the connection

    def done(self):
        return True
    def topic(self):
        return self._name 

# TODO: one RTD manager for all notebooks may have conflicts, but RTD manager must
# be created on main thread.
_rtdmgr = xlo.RtdManager()

class JupyterConnection:
    
    _pending_messages = dict() # str -> Future

    def __init__(self, connection_file, xloil_path):

        from jupyter_client.asynchronous import AsyncKernelClient

        self._loop = asyncio.get_event_loop()
        self._client = AsyncKernelClient()
        self._xloil_path = xloil_path.replace('\\', '/')

        cf = jupyter_client.find_connection_file(connection_file)
        xlo.log("Jupyter: found connection for file {connection_file}")
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
            "import xloil_jupyter\n"
            "importlib.reload(xloil_jupyter)\n"
            "_xloil_vars = xloil_jupyter.MonitoredVariables(get_ipython())\n"
        )

        msg = await self._client.get_shell_msg()
        if 'content' in msg:
            content = msg['content']
            if content['status'] == 'error':
                raise Exception(f"Connection failed: {content['evalue']}")

    def invoke(func_name, *args, **kwargs):
        msg = self._client.execute("xloil_ipython._function_invoke("
            f"{func_name}, {_serialise(args)}, {_serialise(kwargs)})")
        future = self._loop.ensure_future() # err?
        self._pending_messages[msg['msg_id']] = future

    def watch_variable(self, name):
        if _rtdmgr.peek(name) is None: # Remove peek, just use dict
            watcher = _VariableWatcher(name, self._client)
            _rtdmgr.start(watcher)
        return _rtdmgr.subscribe(name)

    def publish_variables(self, updates:dict):
        for name, value in updates.items():
            _rtdmgr.publish(name, value)

    async def process_messages(self):
   
        while self._client.is_alive():
            from queue import Empty
            try:
                msg = await self._client.get_iopub_msg()

                if not 'content' in msg: continue
                content = msg['content']

                pending = self._pending_messages.get("parent_msg_id", None)
                if pending is not None:
                    pending.set_result(content)
                    continue

                if not 'data' in content: continue
                data=content['data']

                if not 'xloil/data' in data: continue
                raw_data = data['xloil/data']
                #meta_type = content['metadata']['type']
                xloil_msg = _deserialise(raw_data)

                if type(xloil_msg) is _VariableChangeMessage:
                    self.publish_variables(xloil_msg.updates)
                elif type(xloil_msg) is _FuncRegisterMessage:
                    descr = xloil_msg.description
                    # Re-hack func object, check _func is a str?
                    def stub(*args, **kwargs):
                        self.invoke(descr._func, *args, **kwargs)
                    descr._func = stub
                    xloil_core.register_functions(__module__, [descr.create_holder()]) 
                else:
                    raise Exception(f"Unknown message: {raw_data}")

            except Empty:
                pass


# Todo: there's no way of closing the connection...without using RTD of course! RTD managed resource

# Will definitely need this to avoid double importing functions etc.
_jupyter_connection_cache = dict()


@xlo.func(rtd=True)
async def xloJpyConnect(ConnectionFile):
    #yield "Wait..."
    conn = JupyterConnection(ConnectionFile, os.path.dirname(__file__))
    await conn.connect()
    yield conn
    await conn.process_messages()
    await asyncio.sleep(2)
    yield "Kernel Stopped"

@xlo.func
def xloJpyWatch(Name, Connection):
    return Connection.watch_variable(Name)

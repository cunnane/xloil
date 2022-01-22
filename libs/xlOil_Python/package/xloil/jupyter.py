import xloil as xlo
import xloil.register
import asyncio
import inspect
import IPython
import pickle
from IPython.display import publish_display_data
import jupyter_client
import os
import re

def _serialise(obj):
    return pickle.dumps(obj, protocol=0).decode('latin1')

def _deserialise(dump:str):
    return pickle.loads(dump.encode('latin1'))

class MonitoredVariables:
    """
    Created within the jupyter kernel to hook the 'post_execute' event and watch
    for variable changes
    """
    def __init__(self, ipy_shell):
        # Takes a reference to the ipython shell (e.g. from get_ipython())

        self._values = dict()
        self._shell = ipy_shell

        # Hook post_execute
        ipy_shell.events.register('post_execute', self.post_execute)
    
    def post_execute(self):
        updates = {}
        # Loop through all global variables looking for changes
        for name, val in self._values.items():
            that_val = self._shell.user_ns.get(name, None)
            # Use is to check for equality rather than == as the latter
            # may not return a single value e.g. numpy arrays
            if not that_val is val:
                updates[name] = that_val
                self._values[name] = that_val

        if len(updates) > 0:
            publish_display_data(
                { "xloil/data": _serialise(updates) },
                { 'type': "VariableChange" }
            )

    def watch(self, name):
        # Starts monitoring the given variable name
        self._values[name] = globals().get(name, None)
        # Run the hook now to publish the variable
        self.post_execute()

    def stop_watch(self, name):
        # Stops monitoring the given variable name
        del self._values[name]

    def close(self):
        self._shell.events.unregister('post_execute', self.post_execute)


class _FuncDescription:
    """
        A serialisable func description we can pickle and send over Jupyter 
        messaging
    """
    def __init__(self, func, name, help, override_args):
        self.func_name = func.__name__
        self.name = name
        self.help = help
        self.args, self.return_type = xloil.register.function_arg_info(func)

        self.args = xlo.Arg.override_arglist(self.args, override_args)

    def register(self, connection):
        
        import xloil_core

        func_name = self.func_name

        @xloil.register.async_wrapper
        async def shim(*args, **kwargs):
            return await connection.invoke(func_name, *args, **kwargs)

        core_argspec = [arg.to_core_argspec() for arg in self.args]

        # TODO: can we avoid using the core?
        spec = xloil_core.FuncSpec(shim, 
            args = core_argspec,
            name = self.name,
            features = "rtd",
            help = self.help,
            category = "xlOil Jupyter")

        if self.return_type is not inspect._empty:
            spec.return_converter = xloil.register.find_return_converter(return_type)

        xlo.log(f"Found func: '{str(spec)}'", level="debug")

        # TODO: can we avoid using the core?
        xloil_core.register_functions([spec])

def _replacement_func_decorator(
        fn=None,
        name=None, 
        help="", 
        args=None):

    """
        Replaces xlo.func in jupyter but removes arguments which do not make sense
        when called from jupyter
    """

    def decorate(fn):
       
        spec = _FuncDescription(fn, name or fn.__name__, help, args)

        publish_display_data(
            { "xloil/data": _serialise(spec) },
            { 'type': "FuncRegister" }
        )

        return fn

    return decorate if fn is None else decorate(fn)

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

    """ 
    Rtd publisher which monitors a single variable
    """

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

_rtd_server_object = None
# We use the ipython kernel json filename as the object cache key for the 
# connection object. To preserve this we need to provide a custom converter
# otherwise RtdServer will resolve the cache reference when converting to 
# a python object
_uncached_convert = None

def _rtd_server():
    global _rtd_server_object, _uncached_convert
    if _rtd_server_object is None:
        import xloil.type_converters
        _rtd_server_object = xlo.RtdServer()
        _uncached_convert = xloil.type_converters.get_converter("object", read=True, cache=False)
    return _rtd_server_object

class _JupyterConnection:
    
    _pending_messages = dict() # Dict[str -> Future]
    _watched_variables = dict()
    _registered_funcs = set()
    _ready = False # Indicates whether the connection can receive commands  

    def __init__(self, connection_file, xloil_path):

        from jupyter_client.asynchronous import AsyncKernelClient

        self._loop = xlo.get_async_loop()
        self._client = AsyncKernelClient()
        self._xloil_path = xloil_path.replace('\\', '/')
        self._connection_file = connection_file

        cf = jupyter_client.find_connection_file(self._connection_file)
        xlo.log(f"Jupyter: found connection for file {self._connection_file}", level='debug')
        self._client.load_connection_file(cf)

    async def connect(self):
        
        from queue import Empty
        
        self._client.wait_for_ready(timeout=3)
        # TODO: what if we timeout?

        # Flush shell channel (wait_for_ready doesn't do this for some reason), but we need it 
        # to receive the shell message below
        while True:    
            try:
                msg = await self._client.get_shell_msg(timeout=0.2)
            except Empty:
                break

        # Setup the python path to find xloil.jupyter. We do a reload of xloil because
        # we override xloil.func - this allows repeated connection to the same kernel
        # without error.  We also connect the variable monitor

        # TODO: if the target notebook already has imported xloil under a different name,
        # I suspect the overwrite of xloil.func will not work.
        
        xlo.log(f"Initialising Jupyter connection {self._connection_file}", level='debug')
        
        msg_id = self._client.execute(
            "import sys, IPython\n" + 
            f"if not '{self._xloil_path}' in sys.path: sys.path.append('{self._xloil_path}')\n" + 
            "import xloil\n"
            "import xloil.jupyter\n"
            "xloil.func = xloil.jupyter._replacement_func_decorator\n" # Overide xloil.func to do our own thing
            "_xloil_vars = xloil.jupyter.MonitoredVariables(get_ipython())\n"
        )

        # TODO: retry?
        # Some examples online suggest get_shell_msg(msg_id) should work. It doesn't, which is 
        # a shame as it would be rather nicer than this loop.
        msg = None
        while True:
            if not await self._client.is_alive():
                raise Exception("Jupyter client died")

            try:
                msg = await self._client.get_shell_msg(timeout=1)
            except Empty:
                xlo.log("Waiting for Jupyter initialisation", level='debug')
                continue

            if msg.get('parent_header', {}).get('msg_id', None) == msg_id:
                if msg['content']['status'] == 'error':
                    trace = "\n".join(msg['content']['traceback'])
                    ansi_escape = re.compile(r'(\x9B|\x1B\[)[0-?]*[ -\/]*[@-~]')
                    raise Exception(f"Connection failed: {msg['content']['evalue']} at {ansi_escape.sub('', trace)}")
                break

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
            _rtd_server().drop(topic)
        
        # Remove any stragglers
        self._watched_variables.clear()

        # Remove all registered functions
        xlo.deregister_functions(None, self._registered_funcs)


    async def wait_for_restart(self):

        session = self._sessionId

        while True:
            # We'd expect to see exec_state == 'starting', on iopub but sometimes it doesn't 
            # appear (???) so we settle for any exec_state with a new session id.
            msg = await self._client.get_iopub_msg()
            exec_state = msg.get('content', {}).get('execution_state', None)
            session = msg['header']['session']
            if exec_state is not None and session != self._sessionId:
                break
            await asyncio.sleep(1)

        xlo.log(f"Jupyter restart: new session {session}", level='info')

        await self.connect()

    def execute(self, *args, **kwargs):
        if not self._ready:
            raise JupyterNotReadyError()
        return self._client.execute(*args, **kwargs)

    async def aexecute(self, command):
        """
        Async run the given command on the kernel
        """
        if not self._ready:
            raise JupyterNotReadyError()

        msg_id = self.execute(command)
        future = self._loop.create_future()
        self._pending_messages[msg_id] = future
        return await future

    async def invoke(self, func_name, *args, **kwargs):
        if not self._ready:
            raise JupyterNotReadyError()

        args_data = repr(_serialise(args))
        kwargs_data = repr(_serialise(kwargs))
        msg_id = self.execute("xloil.jupyter.function_invoke("
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

            xlo.log(f"Starting variable watch {name}", level='debug')
            _rtd_server().start(watcher)

        return _rtd_server().subscribe(topic, _uncached_convert)

    def stop_watch_variable(self, name):
        try:
            del self._watched_variables[name]
        except KeyError:
            pass

    def publish_variables(self, updates:dict):
        for name, value in updates.items():
            _rtd_server().publish(self._watch_prefix(name), value)

    async def process_messages(self):
   
        while await self._client.is_alive():
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

            msg_type = msg['header']['msg_type']

            # If kernel is shutting down, break out of loop
            if msg_type == 'shutdown_reply':
                xlo.log(f"Jupyter kernel shutdown: {self._connection_file}", level='info')
                self._ready = False
                return content['restart']

            # Check if this is the reply to one of our pending messages
            parent_id = msg.get('parent_header', {}).get('msg_id', None)
            pending = self._pending_messages.get(parent_id, None)

            # Look for xlOil specific message content
            data = content.get('data', {})
            xloil_data = data.get('xloil/data', None)
            payload = None if xloil_data is None else _deserialise(xloil_data)

            
            # If we matched a pending message, check for an error or a result then
            # publish the result and pop the pending message. We also will also 
            # match kernel status messages which are replies to our execute request
            # so in this case we just continue
            if pending is not None:
                xlo.log(f"Jupyter matched reply to: {parent_id}", level='trace')
                result = None
                if msg_type == 'error':
                    result = content['evalue']
                elif msg_type == 'execute_result':
                    if xloil_data is None:
                        result = eval(content['data']['text/plain'])
                    elif content['metadata']['type'] == "FuncResult":
                        result = payload
                else: # Kernel status and other messages
                    continue
                pending.set_result(result or f"Kernel reply {content} not understood")
                self._pending_messages.pop(parent_id)
                continue

            if xloil_data is None: 
                continue

            meta_type = content['metadata']['type']
            
            xlo.log(f"Jupyter xlOil Msg: {payload}", level='trace')

            # Handle xlOil messages
            if meta_type == "VariableChange":
                self.publish_variables(payload)

            elif meta_type == "FuncRegister":

                payload.register(self)

                # Keep track of our funtions for a clean shutdown
                self._registered_funcs.add(payload.name)

            elif meta_type == "FuncResult":
                xlo.log(f"Unexpected function result: {msg}")
              
            else:
                raise Exception(f"Unknown xlOil message: {meta_type}, {payload}")

    # TODO: support 'with'?
    #def __enter__(self):
    #def __exit__(self, exc_type, exc_value, traceback)


class _JupyterTopic(xlo.RtdPublisher):

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
                        self._cacheRef = xlo.cache.add(conn, key=self._topic)
                        _rtd_server().publish(self._topic, self._cacheRef)
                        restart = await conn.process_messages()
                        conn.close()
                        if not restart:
                            break
                        await conn.wait_for_restart()

                    _rtd_server().publish(self._topic, "KernelShutdown")

                except Exception as e:
                    _rtd_server().publish(self._topic, e)

            self._task = conn._loop.create_task(run())

    def disconnect(self, num_subscribers):
        if num_subscribers == 0:
            self.stop()

            # Cleanup our cache reference
            if self._cacheRef is not None:
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


def _find_kernel_for_notebook(server, token, filename):
    import requests
    r = requests.get(f"{server}/api/sessions?token={token}", timeout=1)
    if r.ok:
        for session in r.json():
            if filename in session.get('notebook', {}).get('path', ""):
                kernel_id = session.get('kernel', {}).get('id', "")
                return f"kernel-{kernel_id}.json"
    return None

def _find_connection_for_notebook(filename):
    import notebook.notebookapp
    for server in notebook.notebookapp.list_running_servers():
        # Note drop trailing slash on server url
        kernel = _find_kernel_for_notebook(server['url'][:-1], server['token'], filename)
        if kernel: return kernel    
    return None


@xlo.func(
    help="Connects to a jupyter (ipython) kernel. Functions created in the kernel "
         "and decorated with xlo.func will be registered with Excel.",
    args={
            'ConnectInfo': "A file of the form 'kernel-XXX.json' (from %connect_info magic) "
                           "or a URL containing a ipynb file"
         }
)
def xloJpyConnect(ConnectInfo):
    from urllib.parse import urlparse

    connection_file = None
    if re.match('kernel-[a-z0-9\-]*.json', ConnectInfo, re.IGNORECASE):
        # Kernel connection file provided
        connection_file = ConnectInfo
    else:
        # Parse as a URL
        url = urlparse(ConnectInfo)
        filename = os.path.split(url.path)[1]
        if os.path.splitext(filename)[1].lower() == '.ipynb':
            # If there's a token, no need to search
            if 'token' in url.query:
                connection_file = _find_kernel_for_notebook(
                    f"{url.scheme}://{url.netloc}", url.query[6:], filename)
            else: # Search through all local jupyter instances
                connection_file = _find_connection_for_notebook(filename)

    # TODO: load notebook if not already loaded.  Maybe even load jupyter if not loaded?

    if connection_file is None:
        raise Exception(f"Could not find connection for {ConnectInfo}")

    topic = connection_file.lower()
    if _rtd_server().peek(topic) is None:
        conn = _JupyterTopic(topic, connection_file, 
                             os.path.join(os.path.dirname(__file__), os.pardir))
        _rtd_server().start(conn)
    return _rtd_server().subscribe(topic, _uncached_convert)


@xlo.func(
    help="Fetches the value of the specifed variable from the given jupyter"
         "connection. Updates it live using RTD.",
    args={
        'Connection': 'Connection ref output from xloJpyConnect',
        'Name': 'Case-sensitive name of a global variable in the jupyter kernel'
    }
)
def xloJpyWatch(Connection, Name):
    if not isinstance(Connection, _JupyterConnection):
        return "Expected a Jpy Connection"
    return Connection.watch_variable(Name)


@xlo.func(
    help="Runs the given command in the connected kernel, i.e. runs "
         "command.format(repr(Arg1), ...)",
    args={
        'Connection': 'Connection ref output from xloJpyConnect',
        'Command': 'A format string which is to be executed',
    }
)
async def xloJpyRun(Connection, Command:str, 
                    Arg1=None, Arg2=None, Arg3=None, Arg4=None, 
                    Arg5=None, Arg6=None, Arg7=None, Arg8=None, 
                    Arg9=None, Arg10=None, Arg11=None, Arg12=None):
    future = Connection.aexecute(
        Command.format(repr(Arg1), repr(Arg2), repr(Arg3), repr(Arg4), 
                       repr(Arg5), repr(Arg6), repr(Arg7), repr(Arg8), 
                       repr(Arg9), repr(Arg10), repr(Arg11), repr(Arg12)))
    return await future
import xloil as xlo

import asyncio
import IPython
import pickle
from IPython.display import publish_display_data
import jupyter_client
import os
import re
import json
import pydoc

from . import func_inspect

from .jupyter_kernel import _xlOilJupyterImpl
_unpickle = _xlOilJupyterImpl._unpickle
_pickle = _xlOilJupyterImpl._pickle
_FuncDescription = _xlOilJupyterImpl._FuncDescription

def _remove_ansi_escapes(s):
    # Escape codes for colours can appear in some jupyter output
    ansi_escape = re.compile(r'(\x9B|\x1B\[)[0-?]*[ -\/]*[@-~]')
    return ansi_escape.sub('', s)


class _FuncDescriptionDecoder(json.JSONDecoder):
    """
    Deserialises _FuncDescription objects.  Handles 3 cases:
        * Dict describing a _FuncDescription
        * Dict describing a func_inspect.Arg
        * Fully qualified type name as string
    """

    def __init__(self, *args, **kwargs):
        json.JSONDecoder.__init__(self, object_hook=self.object_hook, *args, **kwargs)

    def object_hook(self, dct):

        if 'typeof' in dct:
            if isinstance(dct['typeof'], str):
                dct['typeof'] = pydoc.locate(dct['typeof'])
            try:
                return func_inspect.Arg(**dct)
            except TypeError:
                ...

        try:
            return _FuncDescription(**dct)
        except TypeError:
            ...

        return dct

def _register_func_description(desc: _FuncDescription, connection):
        
    import xloil_core
    from xloil.register import async_wrapper, arg_to_funcarg, find_return_converter
    func_name = desc.func_name

    @async_wrapper
    async def shim(*args, **kwargs):
        return await connection.invoke(func_name, *args, **kwargs)

    core_funcargs = [arg_to_funcarg(arg) for arg in desc.args]

    spec = xloil_core._FuncSpec(shim, 
        args = core_funcargs,
        name = desc.name,
        features = "rtd",
        help = desc.help,
        category = "xlOil Jupyter")

    if desc.return_type is not None:
        spec.return_converter = find_return_converter(return_type)

    xlo.log(f"Jupyter registering func: '{str(spec)}'", level="debug")

    xloil_core._register_functions([spec])


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
        # Although we should only need to connect for the first subscriber,
        # if things somehow go out of sync its safer to just make another
        # call to the variable watcher 
        #
        # Given we refer to it several times, you'd think that making '_xloil_jpy_impl'
        # a module level global. But NO: the deletion order during teardown means
        # it's not available when the RTD server wants to tidy up
        #
        self._connection.execute(f"_xloil_jpy_impl._vars.watch('{self._name}')", silent=True)

    def disconnect(self, num_subscribers):
        if num_subscribers == 0:
            self._connection.execute(f"_xloil_jpy_impl._vars.stop_watch('{self._name}')", silent=True)
            return True # Schedule topic for destruction

    def stop(self):
        self._connection.stop_watch_variable(self._name)

    def done(self):
        return True

    def topic(self):
        return self._topic

def _file_to_string(filepath, indent="  "):
    """
        Write a file to string with the given indent per line. 
        Used to send code to the jupyter kernel
    """

    result = ""
    with open(filepath, "r") as file:
        for line in file:
            result += indent + line + '\n'

    return result

class JupyterNotReadyError(Exception):
    pass

_rtd_server_object = None

def _rtd_server():
    global _rtd_server_object
    if _rtd_server_object is None:
        from xloil.type_converters import get_converter
        _rtd_server_object = xlo.RtdServer()
    return _rtd_server_object

class _JupyterConnection:
    
    _pending_messages = dict() # Dict[str -> Future]
    _watched_variables = dict()
    _registered_funcs = set()
    _ready = False # Indicates whether the connection can receive commands  

    def __init__(self, connection_file):

        from jupyter_client.asynchronous import AsyncKernelClient

        self._loop = xlo.get_async_loop()
        self._client = AsyncKernelClient()
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

        #
        # To avoid the remote kernel needing to export or even know about xlOil,
        # we send a small class by concatenating the files jupyter_kernel.py and
        # func_inspect.py. We then instantiate that class and assign it to the 
        # variable `xloil` to give consistent syntax for function declaration.
        # If xloil is actually imported in the kernel, it will overwrite the symbol
        # and some magic in __init__.py should fix up the names.
        #
        
        xlo.log(f"Initialising Jupyter connection {self._connection_file}", level='debug')
        
        excel_hwnd = xlo.excel_state().hwnd
        our_dir = os.path.dirname(os.path.realpath(__file__))

        msg_id = self._client.execute(
            _file_to_string(os.path.join(our_dir, "jupyter_kernel.py"), indent="")
            + _file_to_string(os.path.join(our_dir, "func_inspect.py"), indent="    ")
            + f"\n_xloil_jpy_impl = _xlOilJupyterImpl(get_ipython(), {excel_hwnd})"
            +  "\nimport sys"
            +  "\nif 'xloil' in sys.modules:"
            + f"\n    xloil.func = _xloil_jpy_impl.func"
            + f"\n    xloil.app = _xloil_jpy_impl.app"
            +  "\nelse:"
            + f"\n    xloil = _xloil_jpy_impl"
        )

        # TODO: retry?
        # Some examples online suggest get_shell_msg(msg_id) should work. It doesn't, which is 
        # a shame.  It may be possible to use the undocumented _async_recv_reply function instead.
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
                    trace = _remove_ansi_escapes("\n".join(msg['content']['traceback']))
                    raise Exception(f"Connection failed: {msg['content']['evalue']} at {trace}")
                break

        self._sessionId = msg['header']['session']
        self._watched_variables.clear()
        self._ready = True


    def close(self):

        # If still ready, i.e. not triggered by a kernel restart, clean up our variables
        if self._ready:
            self.execute("_xloil_jpy_impl._vars.unhook()")

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
        """
        Executes asynchronously on the kernel but does not return. Can be
        called from any thread
        """
        if not self._ready:
            raise JupyterNotReadyError()

        async def coro():
            self._client.execute(*args, **kwargs)

        import asyncio
        asyncio.run_coroutine_threadsafe(coro(), self._loop)

    async def aexecute(self, command:str, **kwargs):
        """
        Async run the given command on the kernel. Must be called on same thread 
        as our xloil event loop, e.g. from an RTD function
        """
        if not self._ready:
            raise JupyterNotReadyError()

        # TODO: maybe able to use the undocumented _async_recv_reply 
        # method on the client rather than rolling our own

        msg_id = self._client.execute(command, **kwargs)
        future = self._loop.create_future()
        self._pending_messages[msg_id] = future
        return await future

    async def invoke(self, func_name, *args, **kwargs):
        """
        Must be called on same thread as our xloil event loop, e.g. from an RTD function
        """
        if not self._ready:
            raise JupyterNotReadyError()

        # TODO: won't work with cellerror, need to convert that to None or string or?

        args_data = repr(_pickle(args))
        kwargs_data = repr(_pickle(kwargs))

        return await self.aexecute(
            f"_xloil_jpy_impl._function_invoke("
            f"{func_name}, {args_data}, {kwargs_data})"
            )

    def _watch_prefix(self, name):
        # Just come up with some unique ID for the RTD topic...
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

        return _rtd_server().subscribe(topic)

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
                # We communicate over the public iopub channel 
                # TODO: consider using the private comm channel rather
                msg = await self._client.get_iopub_msg()
                content = msg['content']
            except (Empty, KeyError):
                # Timed out waiting for messages, or msg had no content
                continue 
            
            xlo.log(f"Jupyter Message: {msg['header']}", level='trace')

            msg_type = msg['header']['msg_type']

            # If kernel is shutting down, break out of loop by returning
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
            meta_type = content.get('metadata', {}).get('type', None)

            # If we matched a pending message, check for an error or a result then
            # publish the result and pop the pending message. We also will also 
            # match kernel status messages which are replies to our execute request
            # so in this case we just continue
            if pending is not None:
                xlo.log(f"Jupyter matched '{msg_type}' reply to: {parent_id}", level='trace')
                result = None
                if msg_type == 'error':
                    result = content['evalue']
                elif msg_type == 'display_data':
                    if xloil_data is not None:
                        result = self._process_xloil_message(meta_type, xloil_data, pending)
                    else:
                        result = f"Unexpected result: {data}"
                elif msg_type == 'execute_result':
                    result = eval(data.get('text/plain', "MISSING_RESULT"))
                else: 
                    continue # Kernel status and other messages

                xlo.log(f"Jupyter result for request {parent_id}: {result}", level='trace')
                pending.set_result(result)
                self._pending_messages.pop(parent_id)
                continue

            if xloil_data is None: 
                continue

            self._process_xloil_message(meta_type, xloil_data)
        
    def _process_xloil_message(self, message_type, payload, pending=None):

        if message_type == "VariableChange":
            payload = _unpickle(payload)
            self.publish_variables(payload)

        elif message_type == "FuncRegister":
            # Registrations are sent using JSON serialisation rather than pickle
            # because the _FuncDescription is not declard in the same python
            # module on both sides 
            func_descr = json.loads(payload, cls=_FuncDescriptionDecoder)
            _register_func_description(func_descr, connection=self)
                
            # Keep track of our funtions for a clean shutdown
            self._registered_funcs.add(func_descr.name)

        elif message_type == "FuncResult":
            payload = _unpickle(payload)
            if pending:
                return payload
            else:
                xlo.log(f"Unexpected function result: {payload}")
        else:
            raise Exception(f"Unknown xlOil message: {message_type}")
            
        xlo.log(f"Jupyter xlOil Msg: {payload}", level='trace')


    # TODO: support 'with'?
    #def __enter__(self):
    #def __exit__(self, exc_type, exc_value, traceback)

_connections = dict()

class _JupyterTopic(xlo.RtdPublisher):

    def __init__(self, topic, connection_file): 
        super().__init__()
        self._connection = _JupyterConnection(connection_file)
        self._topic = topic
        self._task = None 

    def connect(self, num_subscribers):

        if self.done():
            conn = self._connection

            async def run():
                try:
                    xlo.log(f"Awaiting connection to Jupyter kernel {self._topic}", level="debug")
                    await conn.connect()
                    while True:
                        _connections[self._topic] = conn
                        _rtd_server().publish(self._topic, self._topic)
                        restart = await conn.process_messages()
                        xlo.log(f"Closing connection to Jupyter kernel {self._topic}", level="debug")
                        conn.close()
                        if not restart:
                            break
                        await conn.wait_for_restart()
                        xlo.log(f"Restarted connection to Jupyter kernel {self._topic}", level="debug")

                    _rtd_server().publish(self._topic, "KernelShutdown")

                except Exception as e:
                    _rtd_server().publish(self._topic, e)
                finally:
                    conn.close()

            self._task = conn._loop.create_task(run())

    def disconnect(self, num_subscribers):
        if num_subscribers == 0:
            self.stop()

            # Cleanup our cache reference
            del _connections[self._topic]

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


@xlo.converter(_JupyterConnection, register=False)
class JupyterConnectionLookup:
    def read(self, x):
        if isinstance(x, str):
            conn = _connections.get(x, None)
            if conn is not None:
                return conn
        raise Exception(f"Expected jupyter connection ref, got '{x}'")


@xlo.func(
    help="Connects to a jupyter (ipython) kernel. Functions created in the kernel "
         "and decorated with xlo.func will be registered with Excel.",
    args={
            'ConnectInfo': "A file of the form 'kernel-XXX.json' (from %connect_info magic) "
                           "or a URL containing a ipynb file"
         }
)
def xloJpyConnect(ConnectInfo: str):
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
        conn = _JupyterTopic(topic, connection_file)
        _rtd_server().start(conn)
    return _rtd_server().subscribe(topic)


@xlo.func(
    help="Fetches the value of the specifed variable from the given jupyter"
         "connection. Updates it live using RTD.",
    args={
        'Connection': 'Connection ref output from xloJpyConnect',
        'Name': 'Case-sensitive name of a global variable in the jupyter kernel'
    }
)
def xloJpyWatch(Connection: JupyterConnectionLookup, Name):
    return Connection.watch_variable(Name)


@xlo.func(
    help="Runs the given command in the connected kernel, i.e. runs "
         "command.format(repr(Arg1), ...)",
    args={
        'Connection': 'Connection ref output from xloJpyConnect',
        'Command': 'A format string which is to be executed',
    }
)
async def xloJpyRun(Connection:JupyterConnectionLookup, 
                    Command:str, 
                    Arg1=None, Arg2=None, Arg3=None, Arg4=None, 
                    Arg5=None, Arg6=None, Arg7=None, Arg8=None, 
                    Arg9=None, Arg10=None, Arg11=None, Arg12=None):
    future = Connection.aexecute(
        Command.format(repr(Arg1), repr(Arg2), repr(Arg3), repr(Arg4), 
                       repr(Arg5), repr(Arg6), repr(Arg7), repr(Arg8), 
                       repr(Arg9), repr(Arg10), repr(Arg11), repr(Arg12)))
    return await future
# This file is sent directly to the jupyter kernel so avoid adding
# imports etc as they will end up in the kernel's global namespace

class _xlOilJupyterImpl:

    @staticmethod
    def _pickle(obj):
        import pickle
        return pickle.dumps(obj, protocol=0).decode('latin1')

    @staticmethod
    def _unpickle(dump:str):
        import pickle
        return pickle.loads(dump.encode('latin1'))

    @classmethod
    def _serialise(cls, obj):
        # Simple json serialiser: serialises the class' dict, skipping 
        # the special Arg._EMPTY type
        import json
        def f(x):
            return { k: v for k, v in x.__dict__.items() if v is not cls.Arg._EMPTY }
        return json.dumps(obj, default=f)

    class _MonitoredVariables:
        """
        Created within the jupyter kernel to hook the 'post_execute' event and watch
        for variable changes
        """
        def __init__(self, ipy_shell):

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
                from IPython.display import publish_display_data
                publish_display_data(
                    { "xloil/data": _xlOilJupyterImpl._pickle(updates) },
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

        def unhook(self):
            self._shell.events.unregister('post_execute', self.post_execute)

    class _FuncDescription:
        """
            A serialisable func description we can send over Jupyter messaging
        """
        def __init__(self, func_name, name, help, args, return_type):
            self.func_name = func_name
            self.name = name
            self.help = help
            self.args = args
            self.return_type = return_type

    @classmethod
    def _function_invoke(cls, func, args_data, kwargs_data):
        from IPython.display import publish_display_data

        args   = cls._unpickle(args_data)
        kwargs = cls._unpickle(kwargs_data)
        result = func(*args, **kwargs)
        publish_display_data(
            { "xloil/data": cls._pickle(result) },
            { 'type': "FuncResult" }
        )
        #return result # Not used, just in case tho

    def __init__(self, ipy_shell, excel_hwnd):
        # Takes a reference to the ipython shell (e.g. from get_ipython())
        self._excel_hwnd = excel_hwnd
        self._vars = self._MonitoredVariables(ipy_shell)

    def func(self,
            fn=None,
            name=None, 
            help="", 
            args=None):

        """
            Replaces xloil.func in jupyter but removes arguments which do not make sense
            when called from jupyter
        """

        def decorate(fn):
            
            from IPython.display import publish_display_data

            func_args, return_type = self.Arg.full_argspec(fn)
            func_args = self.Arg.override_arglist(func_args, args)

            spec = self._FuncDescription(
                fn.__name__, 
                name or fn.__name__, 
                help, 
                func_args,
                return_type)
        
            publish_display_data(
                { "xloil/data": self._serialise(spec) },
                { 'type': "FuncRegister" }
            )

            return fn

        return decorate if fn is None else decorate(fn)

    def app(self):
        """
            Imports xloil and gives access to the connected Excel Application
        """
        # Importing xloil in the jupyter kernel does some magic, see __init__.py
        import xloil   
        return xloil.Application(hwnd=self._excel_hwnd)

    # IMPORTANT: The last (nonblank) line must be indented as the contents  
    # of func_inspect.py are pasted in here and sent to the kernel. That
    # means we hide all our internals in the _xlOilJupyterImpl class  
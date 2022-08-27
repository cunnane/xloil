class Arg:
    """
    Holds the description of a function argument. Can be used with the `xloil.func`
    decorator to specify the argument description.

    Examples
    --------

    ::

        @xloil.func(args = { 
            'a': xloil.Arg("A", "The First Arg", default=3),
            'b': xloil.Arg("B", "Next Arg",      typeof=double),
        })
        def MyFunc(a, b, c):
            ...

    """

    class _EMPTY:
        """ Indicates the absence of a default argument """
        ...


    def __init__(self, name, help="", typeof=None, default=_EMPTY, is_keywords=False):
        """
        Parameters
        ----------

        name: str
            The name of the argument which appears in Excel's function wizard
        help: str, optional
            Help string to display in the function wizard
        typeof: object, optional
            Selects the type converter used to pass the argument value
        default: object, optional
            A default value to pass if the argument is not specified in Excel
        is_keywords: bool, optional
            Denotes the special kwargs argument. xlOil will expect a two-column array
            in Excel which it will interpret as (key, value) pairs and convert to a
            dictionary. A `**kwargs` argument is auto-detected by xlOil so it is 
            unusual to set this parameter explicitly.
        """

        self.typeof = typeof
        self.name = str(name)
        self.help = help
        self.default = default
        self.is_keywords = is_keywords

    @property
    def has_default(self):
        """ 
        Since 'None' is a fairly likely default value, this property indicates 
        whether there was a user-specified default
        """
        return self.default is not self._EMPTY

    @classmethod
    def from_signature(cls, name, param):
        """
        Constructs an `Arg` from a name and an `inspect.param`
        """
        import inspect

        kind = param.kind
        if kind == param.POSITIONAL_ONLY or kind == param.POSITIONAL_OR_KEYWORD:
            arg = cls(name, 
                      default= cls._EMPTY if param.default is inspect._empty else param.default)
            if param.annotation is not param.empty:
                arg.typeof = param.annotation
            return arg

        elif param.kind == param.VAR_KEYWORD: # can type annotions make any sense here?
            return cls(name, is_keywords=True)

        else: 
            raise Exception(f"Unhandled argument '{name}' with type '{kind}'")

    @classmethod
    def full_argspec(cls, func):
        """
        Returns a list of `Arg` for a given function which describe the function's arguments
        """
        import inspect
        sig = inspect.signature(func)
        params = sig.parameters
        args = [cls.from_signature(name, param) for name, param in params.items()]
        ret_type = None if sig.return_annotation is inspect._empty else sig.return_annotation
        return args, ret_type

    @staticmethod
    def override_arglist(arglist, replacements):
        if replacements is None:
            return arglist
        elif not isinstance(replacements, dict):
            replacements = { a.name : a for a in replacements }

        def override_arg(arg):
            override = replacements.get(arg.name, None)
            if override is None:
                return arg
            elif isinstance(override, str):
                arg.help = override
                return arg
            else:
                return override

        return [override_arg(arg) for arg in arglist]

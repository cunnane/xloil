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

    POSITIONAL    = 0
    KEYWORD_ARGS  = 1
    VARIABLE_ARGS = 2

    def __init__(self, name, help="", typeof=None, default=_EMPTY, kind=POSITIONAL):
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
        kind: int, optional
            Denotes the special *args or **kwargs arguments. For kwargs, xlOil will 
            expect a two-column array in Excel which it will interpret as (key, value) 
            pairs and convert to a dictionary. For *args, xlOil adds a large number of 
            extra trailing optional arguments. Both of these are auto-detected by xlOil 
            so it is unusual to set this parameter explicitly.
        """

        self.typeof = typeof
        self.name = str(name)
        self.help = help
        self.default = default
        self.kind = kind

    def __str__(self):
        if self.kind == self.KEYWORD_ARGS:
            return f"**{self.name}"
        elif self.kind == self.VARIABLE_ARGS:
            return f"*{self.name}"
        else:
            default = "=" + str(self.default) if self.has_default else ""
            type_ = getattr(self.typeof, "__name__", self.typeof) if self.typeof else ""
            return f'{self.name}:{type_}{default}'

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
            return cls(name, kind=cls.KEYWORD_ARGS)

        elif param.kind == param.VAR_POSITIONAL:
            return cls(name, kind=cls.VARIABLE_ARGS)

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

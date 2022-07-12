
from ._core import *

from .logging import *

from .register import (
    Arg, 
    func,
    scan_module,
    register_functions,
    Caller,
    )

from .importer import (
    linked_workbook,
    source_addin,
    get_event_loop
    )

from .com import (
    EventsPaused
    )

from .rtd import RtdSimplePublisher

from .type_converters import (
    ExcelValue,
    Cache, 
    SingleValue,
    AllowRange,
    Array,
    converter,
    returner
    )

from .gui import (
    CustomTaskPane,
    find_task_pane,
    create_task_pane
    )


# Small hacky workaround for the jupyter connection feature.
# If we're imported in an ipython kernel and the variable 
# _xloil_jpy_impl exists, it means xloil has connected to tht
# kernel. In which case, fix up the `func` and `app` functions
# to do the same as before
try:
    ipy = get_ipython()
    
    impl = ipy.user_ns.get("_xloil_jpy_impl", None)
    if impl is not None:
        global func, app
        func = impl.func
        app = impl.app
except NameError:
    ... # Not in an ipython kernel
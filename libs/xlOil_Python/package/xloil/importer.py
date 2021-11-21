"""
Importing this module hooks the import and reload functions
"""

import importlib
import importlib.util
import builtins
import sys
import os

from .register import scan_module, _clear_pending_registrations

_real_builtin_import = builtins.__import__

def _import_hook(name, *args, **kwargs):

    module = _real_builtin_import(name, *args, **kwargs)

    # If name is of the form "foo.bar", the result of _real_builtin_import will point to 
    # the top level package, not the module we want to scan
    real_module = sys.modules.get(name, module)
    scan_module(real_module)

    return module

builtins.__import__ = _import_hook


_real_importlib_reload =  importlib.reload

def _reload_hook(*args, **kwargs):
    _clear_pending_registrations(args[0])
    module = _real_importlib_reload(*args, **kwargs)
    scan_module(module)
    return module

importlib.reload = _reload_hook

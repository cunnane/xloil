"""IPython extension to reload modules before executing user code.

Caveats
=======

Reloading Python modules in a reliable way is in general difficult,
and unexpected things may occur. ``%autoreload`` tries to work around
common pitfalls by replacing function code objects and parts of
classes previously in the module with new versions. This makes the
following things to work:

- Functions and classes imported via 'from xxx import foo' are upgraded
  to new versions when 'xxx' is reloaded.

- Methods and properties of classes are upgraded on reload, so that
  calling 'c.foo()' on an object 'c' created before the reload causes
  the new code for 'foo' to be executed.

Some of the known remaining caveats are:

- Replacing code objects does not always succeed: changing a @property
  in a class to an ordinary method or a method to a member variable
  can cause problems (but in old objects only).

- Functions that are removed (eg. via monkey-patching) from a module
  before it is reloaded are not upgraded.

- C extension modules cannot be reloaded, and so cannot be autoreloaded.

- While comparing Enum and Flag, the 'is' Identity Operator is used (even in the case '==' has been used (Similar to the 'None' keyword)).

- Reloading a module, or importing the same module by a different name, creates new Enums. These may look the same, but are not.
"""

# ------------------------------------------------------------------------------
# superreload
# ------------------------------------------------------------------------------
# -----------------------------------------------------------------------------
#  Copyright (C) 2000 Thomas Heller
#  Copyright (C) 2008 Pauli Virtanen <pav@iki.fi>
#  Copyright (C) 2012 The IPython Development Team
#
#  Distributed under the terms of the BSD License.  The full license is in
#  the file COPYING, distributed as part of this software.
# -----------------------------------------------------------------------------
#
# This IPython module is written by Pauli Virtanen, based on the autoreload
# code by Thomas Heller.
#

import sys
import types
import weakref
import gc
from importlib import reload

func_attrs = [
    "__code__",
    "__defaults__",
    "__doc__",
    "__closure__",
    "__globals__",
    "__dict__",
]


def update_function(old, new):
    """Upgrade the code object of a function"""
    for name in func_attrs:
        try:
            setattr(old, name, getattr(new, name))
        except (AttributeError, TypeError):
            pass


def update_instances(old, new):
    """Use garbage collector to find all instances that refer to the old
    class definition and update their __class__ to point to the new class
    definition"""

    refs = gc.get_referrers(old)

    for ref in refs:
        if type(ref) is old:
            object.__setattr__(ref, "__class__", new)


def update_class(old, new):
    """Replace stuff in the __dict__ of a class, and upgrade
    method code objects, and add new methods, if any"""
    for key in list(old.__dict__.keys()):
        old_obj = getattr(old, key)
        try:
            new_obj = getattr(new, key)
            # explicitly checking that comparison returns True to handle
            # cases where `==` doesn't return a boolean.
            if (old_obj == new_obj) is True:
                continue
        except AttributeError:
            # obsolete attribute: remove it
            try:
                delattr(old, key)
            except (AttributeError, TypeError):
                pass
            continue
        except ValueError:
            # can't compare nested structures containing
            # numpy arrays using `==`
            pass

        if update_generic(old_obj, new_obj):
            continue

        try:
            setattr(old, key, getattr(new, key))
        except (AttributeError, TypeError):
            pass  # skip non-writable attributes

    for key in list(new.__dict__.keys()):
        if key not in list(old.__dict__.keys()):
            try:
                setattr(old, key, getattr(new, key))
            except (AttributeError, TypeError):
                pass  # skip non-writable attributes

    # update all instances of class
    update_instances(old, new)


def update_property(old, new):
    """Replace get/set/del functions of a property"""
    update_generic(old.fdel, new.fdel)
    update_generic(old.fget, new.fget)
    update_generic(old.fset, new.fset)


def isinstance2(a, b, typ):
    return isinstance(a, typ) and isinstance(b, typ)


UPDATE_RULES = [
    (lambda a, b: isinstance2(a, b, type), update_class),
    (lambda a, b: isinstance2(a, b, types.FunctionType), update_function),
    (lambda a, b: isinstance2(a, b, property), update_property),
]
UPDATE_RULES.extend(
    [
        (
            lambda a, b: isinstance2(a, b, types.MethodType),
            lambda a, b: update_function(a.__func__, b.__func__),
        ),
    ]
)


def update_generic(a, b):
    for type_check, update in UPDATE_RULES:
        if type_check(a, b):
            update(a, b)
            return True
    return False


class StrongRef:
    def __init__(self, obj):
        self.obj = obj

    def __call__(self):
        return self.obj


mod_attrs = [
    "__name__",
    "__doc__",
    "__package__",
    "__loader__",
    "__spec__",
    "__file__",
    "__cached__",
    "__builtins__",
]


def append_obj(module, d, name, obj, autoload=False):
    in_module = hasattr(obj, "__module__") and obj.__module__ == module.__name__
    if autoload:
        # check needed for module global built-ins
        if not in_module and name in mod_attrs:
            return False
    else:
        if not in_module:
            return False

    key = (module.__name__, name)
    try:
        d.setdefault(key, []).append(weakref.ref(obj))
    except TypeError:
        pass
    return True


def superreload(module, reload=reload, old_objects=None, shell=None):
    """Enhanced version of the builtin reload function.

    superreload remembers objects previously in the module, and

    - upgrades the class dictionary of every old class in the module
    - upgrades the code object of every old function and method
    - clears the module's namespace before reloading

    """
    if old_objects is None:
        old_objects = {}

    # collect old objects in the module
    for name, obj in list(module.__dict__.items()):
        if not append_obj(module, old_objects, name, obj):
            continue
        key = (module.__name__, name)
        try:
            old_objects.setdefault(key, []).append(weakref.ref(obj))
        except TypeError:
            pass

    # reload module
    try:
        # clear namespace first from old cruft
        old_dict = module.__dict__.copy()
        old_name = module.__name__
        module.__dict__.clear()
        module.__dict__["__name__"] = old_name
        module.__dict__["__loader__"] = old_dict["__loader__"]
    except (TypeError, AttributeError, KeyError):
        pass

    try:
        module = reload(module)
    except:
        # restore module dictionary on failed reload
        module.__dict__.update(old_dict)
        raise

    # iterate over all objects and update functions & classes
    for name, new_obj in list(module.__dict__.items()):
        key = (module.__name__, name)
        if key not in old_objects:
            # here 'shell' acts both as a flag and as an output var
            if (
                shell is None
                or name == "Enum"
                or not append_obj(module, old_objects, name, new_obj, True)
            ):
                continue
            shell.user_ns[name] = new_obj

        new_refs = []
        for old_ref in old_objects[key]:
            old_obj = old_ref()
            if old_obj is None:
                continue
            new_refs.append(old_ref)
            update_generic(old_obj, new_obj)

        if new_refs:
            old_objects[key] = new_refs
        else:
            del old_objects[key]

    return module


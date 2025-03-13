# We use tomlkit as it preserves comments.
from collections import namedtuple
import tomlkit as toml
import winreg as reg
import xloil
from pathlib import Path
from itertools import islice
import sys
import os

from xloil.stubs.xloil_core import CellError

@xloil.func
def test_name(x):
    return x

@xloil.func
async def test_name2(x):
    yield x

class Settings:
    """
        Manages accessing and saving the settings file
    """
    def __init__(self, path):
        self._doc = toml.parse(Path(path).read_text())
        self._path = path

    def __getitem__(self, *args):
        return self._doc.__getitem__(*args)

    def set_env_var(self, name, value):
        """ Sets the value of a token in the Python plugin's Environment block """
        table = self._find_table(self.python['Environment'], name)
        table[name] = value

    def get_env_var(self, name):
        """ Returns the value of a token in the Python plugin's Environment block """
        table = self._find_table(self.python['Environment'], name)
        return table[name]

    def set_addin_env_var(self, name, value):
        """ Sets the value of a token in the xlOil addin's Environment block """
        table = self._find_table(self.addin['Environment'], name)
        table[name] = value

    def get_addin_env_var(self, name):
        table = self._find_table(self.addin['Environment'], name)
        return table[name]

    def save(self):
        with open(self._path, "w") as file:
            toml.dump(self._doc, file)

    @property
    def python(self):
        return self._doc['xlOil_Python']
    
    @property
    def addin(self):
        return self._doc['Addin']
    
    @property
    def path(self):
        return self._path

    @staticmethod
    def _find_table(array_of_tables, key):
        """
            We may have multiple (ordered) tables with the same name, e.g.
            the enviroment blocks. This function returns the table containing
            the specified key.
        """
        # May be just a single table
        if key in array_of_tables:
            return array_of_tables

        for table in array_of_tables:
            if key in table:
                return table
        raise KeyError(f"No table contains {key}")

_settings = Settings(xloil.source_addin().settings_file)


#
# Log Level Callbacks
# --------------------
# 
async def set_log_level(ctrl, id, index):
    value = xloil.log.levels[index]
    xloil.log.level = value # thread safe?
    _settings['Addin']['LogLevel'] = value
    _settings.save()

def get_log_level_count(ctrl):
    return len(xloil.log.levels)

def get_log_level(ctrl, i):
    return xloil.log.levels[i]

def get_log_level_selected(ctrl):
    return xloil.log.levels.index(xloil.log.level)

#
# Date Format Callbacks
# ---------------------
#
async def set_date_formats(ctrl, value):
    values = value.split(";")
    _settings['Addin']['DateFormats'] = values
    _settings.save()

    # Update the formats currently in use
    xloil.date_formats.clear()
    xloil.date_formats.extend(values)

def get_date_formats(ctrl):
    return ";".join(_settings['Addin']['DateFormats'])

#
# User search path callbacks
# --------------------------
#

async def set_user_search_path(ctrl, value):
    paths = value.split(";")
    for path in paths: 
        if not path in sys.path:
            sys.path.append(path)

    # Save the settings afterwards in case the above fails
    _settings.set_env_var("XLOIL_PYTHON_PATH",";" if len(value) == 0 else value)
    _settings.save()
    
def get_user_search_path(ctrl):
    value = _settings.get_env_var("XLOIL_PYTHON_PATH")
    xloil.log(f"Search Path: {value}")
    return "" if value == ";" else value

#
# PYTHONEXECUTABLE path callbacks
# -------------------------------
#

async def set_python_home(ctrl, value):
    _settings.set_env_var("PYTHONEXECUTABLE", value)
    _settings.save()

    restart_notify()
    
def get_python_home(ctrl):
    return _settings.get_env_var("PYTHONEXECUTABLE")

 
#
# PYTHONPATH callbacks
# --------------------
#

def set_python_path(ctrl, value):
    _settings.set_env_var("PYTHONPATH", value)
    _settings.save()

    restart_notify()

def get_python_path(ctrl):
    return _settings.get_env_var("PYTHONPATH")

#
# Python Environment callbacks
# ----------------------------
#

_PythonEnv = namedtuple("_PythonEnv", "name, version, executable")

def _find_python_enviroments_from_key(pythons_key):

    environments = []
    try:
        i = 0
        while True:
            vendor = reg.EnumKey(pythons_key, i)
            i += 1
            with reg.OpenKey(pythons_key, vendor) as vendor_key:
                    
                try:
                    j = 0
                    while True:
                        version = reg.EnumKey(vendor_key, j)
                        j += 1
                        with reg.OpenKey(vendor_key, version) as kVersion:
                            name = reg.QueryValueEx(kVersion, 'DisplayName')[0]
                            install_path = reg.OpenKey(kVersion, 'InstallPath')
                            environments.append(_PythonEnv(
                                name=name, 
                                version=reg.QueryValueEx(kVersion, 'SysVersion')[0],
                                executable=reg.QueryValueEx(install_path, 'ExecutablePath')[0]))
                except OSError:
                    ...
    except OSError:
        ...

    return environments


def _find_conda_environments():

    from pathlib import Path
    env_file = Path.home() / '.conda' / 'environments.txt'
    if not env_file.exists():
        return []

    env_paths = set(x for x in env_file.read_text().split('\n') if len(x) > 0)

    environments = []

    # TODO: currently not sure of the easiest/best way to get version info
    for path in env_paths:
        root, _, env = path.rpartition(os.path.sep + "envs" + os.path.sep)
        exe_path = os.path.join(path, "python.exe")
        if root in env_paths:
            environments.append(_PythonEnv(
                name=f'{env} ({os.path.basename(root)})',
                version="?",
                executable=exe_path))
        else:
            environments.append(_PythonEnv(
                name=os.path.basename(path),
                version="?",
                executable=exe_path))

    return environments

def _find_env_by_exe(environments, filename):
    filename = filename.upper()
    for i in range(len(environments)):
        if environments[i].executable.upper() == filename.upper():
            return i

    return None


def _find_python_enviroments():

    roots = [reg.HKEY_LOCAL_MACHINE, reg.HKEY_CURRENT_USER]
    environments = []

    for root in roots:
        try:
            with reg.OpenKey(root, "Software\\Python") as pythons_key:
                environments += _find_python_enviroments_from_key(pythons_key)
        except FileNotFoundError:
            ... # Reg key doesn't exist, try the next one

    for env in _find_conda_environments():
        if env.executable not in environments:
            environments.append(env)

    py_exe = _settings.get_env_var("PYTHONEXECUTABLE")

    # Check if current environment is already described 
    if _find_env_by_exe(environments, py_exe) is None:
        environments.append(_PythonEnv(
            name='Current',
            version=f'{sys.version_info.major}.{sys.version_info.minor}',
            executable=py_exe))

    return environments


_PYTHON_ENVIRONMENTS = _find_python_enviroments()


async def set_python_environment(ctrl, id, index):

    environment = _PYTHON_ENVIRONMENTS[index]
    exe_path = environment.executable

    xloil_bin_path = Path(exe_path).parent / "share/xloil"

    if not xloil_bin_path.exists():
        xloil.log.error("Changed target python environment to '%s', but the xlOil package is missing. " +
                        "Unless it is installed, xlOil will not load correctly when Excel is restarted.",
                        exe_path)

    _settings.set_env_var("PYTHONEXECUTABLE", exe_path)

    # Clear the version override if set (it shouldn't generally be required)
    _settings.set_env_var("XLOIL_PYTHON_VERSION", "")

    # This is where we look for the binaries
    try:
        _settings.set_addin_env_var("XLOIL_PATH", str(xloil_bin_path))
    except toml.exceptions.NonExistentKey:
        ...

    _settings.save()

    # Invalidate controls
    _ribbon_ui.invalidate("PYTHONEXECUTABLE")

    restart_notify()


def get_python_environment_count(ctrl):
    return len(_PYTHON_ENVIRONMENTS)


def get_python_environment(ctrl, i):
    return _PYTHON_ENVIRONMENTS[i].name


def get_python_environment_selected(ctrl):
    py_home = _settings.get_env_var("PYTHONEXECUTABLE")
    return _find_env_by_exe(_PYTHON_ENVIRONMENTS, py_home) or (len(_PYTHON_ENVIRONMENTS) - 1)

#
# Python Load Modules callbacks
# -----------------------------
#

async def set_load_modules(ctrl, value):
    # Allow a semi-colon separator
    modules = value.replace(";", ",").split(",")
    _settings.python['LoadModules'] = modules
    _settings.save()

    # Load any new modules. Catch errors in case of misspelt names
    import importlib
    for m in modules:
        try:
            if not m in sys.modules:
                importlib.import_module(m)
        except Exception as err:
            xloil.log.warn(f"Ribbon failed loading module {str(m)}: {str(err)}")

def get_load_modules(ctrl):
    value = _settings.python['LoadModules']
    return ",".join(value)


#
# Python Debugger callbacks
# -------------------------
#

def get_debugger_count(ctrl):
    import xloil.debug
    return len(xloil.debug.DEBUGGERS)

def get_debugger(ctrl, i):
    import xloil.debug
    return xloil.debug.DEBUGGERS[i]

def get_debugger_selected(ctrl):
    import xloil.debug
    current = _settings.python['Debugger']
    i = xloil.debug.DEBUGGERS.index(current)
    return max(i, 0)

async def set_debugger(ctrl, id, index):
    import xloil.debug

    choice = xloil.debug.DEBUGGERS[index]
    _settings.python['Debugger'] = choice
    _settings.save()

    xloil.debug.use_debugger(
        choice, 
        port=int(_settings.python['DebugPyPort']))

def _find_free_port():
    #https://stackoverflow.com/questions/1365265
    import socket
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        return s.getsockname()[1]

def _is_port_in_use(port: int) -> bool:
    # https://stackoverflow.com/questions/2470971/
    import socket
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(('localhost', port))
            return False
    except OSError:
        return True

def get_debugpy_port(ctrl):
    port = int(_settings.python['DebugPyPort'])
    if _is_port_in_use(port):
        port = _find_free_port()
        _settings.python['DebugPyPort'] = port
        # We don't save the setting as the port may be free next time
    return port

async def set_debugpy_port(ctrl, value):
    _settings.python['DebugPyPort'] = value
    _settings.save()


#
# Open Log callback
# -----------------
#

async def press_open_log(ctrl):
    xloil.log.flush()
    import os
    os.startfile(xloil.log.path)

#
# Open Console callbacks
# ----------------------
#
async def press_open_console(ctrl):

    def open_console(root):
        from xloil.gui.tkinter import TkConsole
        import tkinter
        import code

        top_level = tkinter.Toplevel(root)
        console = TkConsole(top_level, code.interact,
            fg='white', bg='black', font='Consolas', insertbackground='red')
        console.pack(expand=True, fill=tkinter.BOTH)
        console.bind("<<CommandDone>>", lambda e: top_level.destroy())

        top_level.deiconify()

    from xloil.gui.tkinter import Tk_thread
    await Tk_thread().submit_async(open_console, Tk_thread().root)


theConsoleQt=None

async def press_open_qtconsole(ctrl):

    def open_console():
        global theConsoleQt
        from xloil.gui.qt_console import create_qtconsole_inprocess
        console = create_qtconsole_inprocess()
        console.show()
        theConsoleQt = console

    from xloil.gui.qtpy import Qt_thread
    await Qt_thread().submit_async(open_console)


#
# Restart Notify callbacks
# ------------------------
#
_restart_label_visible = False

def restart_notify():
    global _restart_label_visible, _ribbon_ui
    _restart_label_visible = True
    _ribbon_ui.invalidate("RestartLabel")

def get_restart_label_visible(ctrl):
    global _restart_label_visible
    return _restart_label_visible



def set_error_propagation(ctrl, value):
    _settings.addin['ErrorPropagation'] = value
    _settings.save()
    # TODO: enable this when the settings update is live otherwise the behaviour is a little unexpected
    #for func in xloil.core_addin().functions():
    #   func.error_propagation = value


def get_error_propagation(ctrl):
    return bool(_settings.addin.get('ErrorPropagation', False))

def _fix_name_errors(ctrl):    
    xloil.fix_name_errors(xloil.active_workbook())


#
# Ribbon creation
# ---------------
#

def _ribbon_func_map(func: str):
    # Just finds the function with the given name in this module 
    xloil.log.debug(f"Calling xlOil Ribbon '{func}'...")
    return globals()[func]

_ribbon_ui = xloil.ExcelGUI(ribbon=r'''
    <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
      <ribbon>
        <tabs>
          <tab id="tab0" label="xlOil Py">
            <group id="grp0" autoScale="false" centerVertically="false" label="Environment">
              <dropDown id="ddc0" label="Environment" 
                screentip="Python environments listed in the Windows Registry"
                supertip="Environment changes only take effect after restarting Excel"
                onAction="set_python_environment"
                getItemCount="get_python_environment_count"
                getItemLabel="get_python_environment"
                getSelectedItemIndex="get_python_environment_selected" />
              <editBox id="PYTHONEXECUTABLE" label="PYTHONEXECUTABLE" 
                screentip="The location of python.exe in the distribution"
                supertip="Environment changes only take effect after restarting Excel"
                sizeString="c:/a/path/is/this/size"
                getText="get_python_home" 
                onChange="set_python_home" />
              <editBox id="PYTHONPATH" label="PYTHONPATH" sizeString="c:/a/path/is/this/size"
                screentip="A semi-colon separated list of module search directories"
                supertip="Prefer to use this for system paths and add user directories to the Search Paths"
                getText="get_python_path" 
                onChange="set_python_path" />
            </group>
            <group id="grp1" autoScale="false" centerVertically="false" label="Modules">
              <editBox id="ebx3" label="Load Modules" sizeString="a module; another module; another"
                screentip="Python modules to load"
                supertip="The modules must be available on python's sys.path. Use a comma to separate."
                getText="get_load_modules" 
                onChange="set_load_modules"/>
              <editBox id="ebx4" label="Search Paths" sizeString="a module; another module; another"
                screentip="Paths added to python's sys.path"
                supertip="Prefer to add user directories here rather than editing PYTHONPATH directly. Use the usual semi-colon (;) separator and the path separator (\)"
                getText="get_user_search_path"
                onChange="set_user_search_path"/>
              <labelControl id="RestartLabel" label="!Restart Required!" 
                getVisible="get_restart_label_visible" />
            </group>
            <group id="grp2" autoScale="false" centerVertically="false" label="Debugging" >
              <button id="btn5" size="large" label="Console" imageMso="ContainerGallery" 
                onAction="press_open_console"/>
              <button id="qtconsole" size="large" label="QtConsole" imageMso="ContainerGallery" 
                onAction="press_open_qtconsole"
                screentip="Opens a Jupyter console to an inprocess kernel"
                supertip="The console is connected to the Excel application, so can register functions and manipulate workbooks"/>
              <button id="btn6" size="large" label="Open Log" imageMso="ZoomCurrent75" 
                onAction="press_open_log"
                screentip="Opens the log file"
                supertip="Uses the editor configured to open '.log' files"/>
              <dropDown id="ddc8" label="Log Level" 
                screentip="The threshold level to control which messages are written to the log file"
                getSelectedItemIndex="get_log_level_selected" 
                getItemCount="get_log_level_count"
                getItemLabel="get_log_level"
                onAction="set_log_level" >
              </dropDown>
              <dropDown id="ddc9" label="Debugger" 
                screentip="Debugger used"
                supertip="Check the documentation for details on debugging"
                getSelectedItemIndex="get_debugger_selected" 
                getItemCount="get_debugger_count"
                getItemLabel="get_debugger"
                onAction="set_debugger" >
              </dropDown>
              <editBox id="ebxDebugPy" label="DebugPy Port"
                screentip="Connection port used if the debugpy/vscode debugger is selected"
                getText="get_debugpy_port"
                onChange="set_debugpy_port"/>
            </group>
            <group id="grp3" autoScale="false" centerVertically="false" label="Other" >
              <editBox id="ebxDateFormats" label="Date Formats" 
                screentip="Date formats tried when parsing strings as dates"
                supertip="Uses 'std::get_time' formats: %Y year, %m month number, %b month name, etc." 
                getText="get_date_formats" 
                onChange="set_date_formats"/>
              <checkBox id="cbxPropagation" label="Propagate Errors"
                screentip="If enabled, any error code (eg. \#NUM!) arguments are passed through as the function's return 
                            value, otherwise all argument values are handled by the function"
                supertip="(requires restart)"
                getPressed="get_error_propagation"              
                onAction="set_error_propagation"/>
              <button id="fixNameErrors" size="normal" label="Fix #NAME!" imageMso="ErrorChecking" 
                onAction="_fix_name_errors"
                screentip="Marks #NAME! errors in current workbook for recalculation."
                supertip="These errors cannot simply be resolved just by Ctrl-Alt-F9. May not be performant in 
                          large workbooks with many errors."/>
            </group>
          </tab>
        </tabs>
      </ribbon>
    </customUI>
    ''', 
    funcmap=_ribbon_func_map)
# We use tomlkit as it preserves comments.
import tomlkit as toml
import winreg as reg
import xloil
from pathlib import Path
from itertools import islice
import sys
import os

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
        table = self._find_table(self.python['Environment'], name)
        table[name] = value

    def get_env_var(self, name):
        table = self._find_table(self.python['Environment'], name)
        return table[name]

    def save(self):
        with open(self._path, "w") as file:
            toml.dump(self._doc, file)

    @property
    def python(self):
        return self._doc['xlOil_Python']

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
# -------------------------
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

def _find_python_enviroments_from_key(pythons_key):

    environments = {}
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
                            environments[name] = {
                                'DisplayName': name,
                                'Version':  reg.QueryValueEx(kVersion, 'SysVersion')[0],
                                'ExecutablePath': reg.QueryValueEx(install_path, 'ExecutablePath')[0]
                            }
                except OSError:
                    ...
    except OSError:
        ...

    return environments

def _find_python_enviroments():

    roots = [reg.HKEY_LOCAL_MACHINE, reg.HKEY_CURRENT_USER]
    environments = {}

    for root in roots:
        try:
            with reg.OpenKey(root, "Software\\Python") as pythons_key:
                environments.update(_find_python_enviroments_from_key(pythons_key))
        except FileNotFoundError:
            ... # Reg key doesn't exist, try the next one

    return environments

_python_enviroments = list(_find_python_enviroments().values())

async def set_python_environment(ctrl, id, index):
    environment = _python_enviroments[index]

    _settings.set_env_var("PYTHONEXECUTABLE", environment['ExecutablePath'])

    # Clear the version override if set (it shouldn't generally be required)
    _settings.set_env_var("XLOIL_PYTHON_VERSION", "")

    _settings.save()

    # Invalidate controls
    _ribbon_ui.invalidate("PYTHONEXECUTABLE")

    restart_notify()

def get_python_environment_count(ctrl):
    py_home = _settings.get_env_var("PYTHONEXECUTABLE").upper()

    # Check if current environment is already described in registry
    for env in _python_enviroments:
        if env['ExecutablePath'].upper() == py_home:
            return len(_python_enviroments)

    _python_enviroments.append({
        'DisplayName': 'Current',
        'Version':  f'{sys.version_info.major}.{sys.version_info.minor}',
        'ExecutablePath': _settings.get_env_var("PYTHONEXECUTABLE")
    })

    return len(_python_enviroments)

def get_python_environment(ctrl, i):
    return _python_enviroments[i]['DisplayName']

def get_python_environment_selected(ctrl):
    py_home = _settings.get_env_var("PYTHONEXECUTABLE").upper()
    for i in range(len(_python_enviroments)):
        if _python_enviroments[i]['ExecutablePath'].upper() == py_home:
            return i

    return 0

#
# Python Load Modules callbacks
# -----------------------------
#

async def set_load_modules(ctrl, value):
    modules = value.split(",")
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
# Open Console callback
# ---------------------
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
                onChange="set_debugpy_port" />
            </group>
            <group id="grp3" autoScale="false" centerVertically="false" label="Other" >
             <editBox id="ebxDateFormats" label="Date Formats" 
                screentip="Date formats tried when parsing strings as dates"
                supertip="Uses 'std::get_time' formats: %Y year, %m month number, %b month name, etc." 
                getText="get_date_formats" 
                onChange="set_date_formats"  />
            </group>
          </tab>
        </tabs>
      </ribbon>
    </customUI>
    ''', 
    funcmap=_ribbon_func_map)
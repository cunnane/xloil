# We use tomlkit as it preserves comments.
import tomlkit as toml
import winreg as reg
import xloil
from pathlib import Path
from itertools import islice
import sys

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
            We can have multiple (ordered) tables with the same name, e.g.
            the enviroment blocks. This function returns the table containing
            the specified key.
        """
        for table in array_of_tables:
            if key in table:
                return table
        raise Exception(f"No table contains {key}")

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
    _settings.set_env_var("XLOIL_PYTHON_PATH", value)
    _settings.save()
    paths = value.split(";")
    for path in paths: 
        if not path in sys.path:
            sys.path.append(path)
    
def get_user_search_path(ctrl):
    xloil.log(f"Search Path: {_settings.get_env_var('XLOIL_PYTHON_PATH')}")
    return _settings.get_env_var("XLOIL_PYTHON_PATH")

#
# PYTHONHOME path callbacks
# -------------------------
#

async def set_python_home(ctrl, value):
    _settings.set_env_var("PYTHONHOME", value)
    _settings.save()

    restart_notify()
    
def get_python_home(ctrl):
    return _settings.get_env_var("PYTHONHOME")

 
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

def _find_python_enviroments():

    environments = {}
    with reg.OpenKey(reg.HKEY_LOCAL_MACHINE, "Software\\Python") as kPythons:
        try:
            i = 0
            while True:
                vendor = reg.EnumKey(kPythons, i)
                i += 1
                with reg.OpenKey(kPythons, vendor) as kVendor:
                    
                    try:
                        j = 0
                        while True:
                            version = reg.EnumKey(kVendor, j)
                            j += 1
                            with reg.OpenKey(kVendor, version) as kVersion:
                                name = reg.QueryValueEx(kVersion, 'DisplayName')[0]
                                environments[name] = {
                                    'DisplayName': name,
                                    'SysVersion':  reg.QueryValueEx(kVersion, 'SysVersion')[0],
                                    'PythonPath':  reg.QueryValue(kVersion,   'PythonPath'),
                                    'InstallPath': reg.QueryValue(kVersion,   'InstallPath')
                                }
                    except OSError:
                        ...
        except OSError:
            ...

    return environments

_python_enviroments = list(_find_python_enviroments().values())

async def set_python_environment(ctrl, id, index):
    environment = _python_enviroments[index]

    _settings.set_env_var("PYTHONPATH", "%PYTHONPATH%;" + environment['PythonPath'])
    _settings.set_env_var("PYTHONHOME", environment['InstallPath'])
    _settings.set_env_var("XLOIL_PYTHON_VERSION", environment['SysVersion'])
    _settings.save()

    # Invalidate controls
    _ribbon_ui.invalidate("PYTHONPATH")
    _ribbon_ui.invalidate("PYTHONHOME")

    restart_notify()

def get_python_environment_count(ctrl):
    py_home = _settings.get_env_var("PYTHONHOME").upper()
    for env in _python_enviroments:
        if env['InstallPath'].upper() == py_home:
            return len(_python_enviroments)

    _python_enviroments.append({
        'DisplayName': 'Current',
        'SysVersion':  f'{sys.version_info.major}.{sys.version_info.minor}',
        'PythonPath':  _settings.get_env_var("PYTHONPATH"),
        'InstallPath': _settings.get_env_var("PYTHONHOME")
    })

    return len(_python_enviroments)

def get_python_environment(ctrl, i):
    return _python_enviroments[i]['DisplayName']

def get_python_environment_selected(ctrl):
    py_home = _settings.get_env_var("PYTHONHOME").upper()
    for i in range(len(_python_enviroments)):
        if _python_enviroments[i]['InstallPath'].upper() == py_home:
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

def get_debugpy_port(ctrl):
    return _settings.python['DebugPyPort']

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

    console_choice = _settings.python['RibbonConsole'].lower()

    if console_choice == "tk":

        from xloil.gui.tkinter import Tk_thread, create_tk_console_window
        await Tk_thread().submit_async(create_tk_console_window, Tk_thread().root)

    elif console_choice == "ipython" or console_choice == "ipythonmain":

        def open_console_qt():
            from xloil.gui.qt_console import create_qt_ipython_console
            console = create_qt_ipython_console(console_choice == "ipythonmain")
            console.show()

        from xloil.gui.qtpy import Qt_thread
        await Qt_thread().submit_async(open_console_qt)


async def press_open_jupyter(ctrl):
    from xloil.jupyter_launcher import open_attached_notebook
    await open_attached_notebook()


def get_image_open_jupyter(ctrl):
    from PIL import Image
    icon_path = Path(sys.prefix) / "Menu/jupyter.ico"
    return Image.open(str(icon_path))


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
              <editBox id="PYTHONHOME" label="PYTHONHOME" 
                screentip="The python distribution root directory"
                supertip="Environment changes only take effect after restarting Excel"
                sizeString="c:/a/path/is/this/size"
                getText="get_python_home" 
                onChange="set_python_home" />
              <editBox id="PYTHONPATH" label="PYTHONPATH" sizeString="c:/a/path/is/this/size"
                screentip="A semi-colon separated list of module search directories"
                supertip="Prefer to add user directories to the user search path and leave PYTHONPATH for system directories"
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
                supertip="Prefer to add user directories here and and leave PYTHONPATH for system directories. Use a semi-colon separator"
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
            <group id="grp4" autoScale="false" centerVertically="false" label="Jupyter" >
                <button id="btnOpenJuputer" size="large" label="Jupyter" getImage="get_image_open_jupyter" 
                    onAction="press_open_jupyter"/>
            </group>
          </tab>
        </tabs>
      </ribbon>
    </customUI>
    ''', 
    funcmap=_ribbon_func_map)
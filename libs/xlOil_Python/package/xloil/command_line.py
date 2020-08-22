import subprocess
import sys
import os
from pathlib import Path
import re

_powershellMessage = "Failed to execute script. You may need to change powershell permissions " + \
                     "by typing 'powershell -Command Set-ExecutionPolicy -Scope CurrentUser RemoteSigned'"

def _runPowerShell(script):
    result = subprocess.run(f"powershell {script}", shell=True)
    if result.returncode != 0:
        raise Exception(_powershellMessage)

def _script_dir():
    return os.path.join(sys.prefix, "share", "xloil")

def _install_xloil():
    target_script = os.path.join(_script_dir(), "xloil_Install.ps1")
    _runPowerShell(target_script)

    # Edit the xloil.ini file. To preserve comments and whitespace it's easier to just do 
    # a regex replace rather than read the file as structured TOML

    ini_path = Path(os.getenv('APPDATA')) / "xlOil" / "xlOil.ini"
    ini_txt = ini_path.read_text()
    
    python_path = os.path.join(sys.prefix, "Lib") + ";" + os.path.join(sys.prefix, "DLLs") 
    python_ver = f'{sys.version_info.major}.{sys.version_info.minor}'

    def toml_lit_string(s):
        return "'''" + s.replace('\\','\\\\') + "'''"

    ini_txt, count1 = re.subn(r'^(\s*PYTHONPATH\s*=).*', r'\g<1>' + toml_lit_string(python_path), ini_txt, flags=re.M)
    ini_txt, count2 = re.subn(r'^(\s*PYTHON_LIB\s*=).*', r'\g<1>' + toml_lit_string(sys.prefix), ini_txt, flags=re.M)
    ini_txt, count3 = re.subn(r'^(\s*xlOilPythonVersion\s*=).*', rf'\g<1>"{python_ver}"', ini_txt, flags=re.M)

    if count1 != 1 or count2 != 1 or count3 != 1:
        print(f'WARNING: Failed to set python paths in {ini_path}. You may have to do this manually.')
    else:
        ini_path.write_text(ini_txt)
        print(f'Edited {ini_path} to point to {sys.prefix} python distribution.')

def _remove_xloil():
    target_script = os.path.join(_script_dir(), "xloil_Remove.ps1")
    _runPowerShell(target_script)

def _create_addin(args):
    if len(args) != 1:
        raise Exception("'create' should have one argument, the target filename")
    target_script = os.path.join(_script_dir(), "xloil_NewAddin.ps1")
    _runPowerShell(f'{target_script} {args[0]}')

def main():
    command = sys.argv[1].lower() if len(sys.argv) > 1 else ""

    if command == 'install':
        _install_xloil()
    elif command == 'remove':
        _remove_xloil()
    elif command == 'uninstall':
        _remove_xloil()
    elif command == 'create':
        _create_addin(sys.argv[2:])
    else:
        raise Exception("Syntax: xloil {install, remove, create}")



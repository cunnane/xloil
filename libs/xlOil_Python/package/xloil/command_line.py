import subprocess
import sys
import os

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



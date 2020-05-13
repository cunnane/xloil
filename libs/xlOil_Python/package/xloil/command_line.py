import subprocess
import sys
import os

def _script_dir():
    return os.path.join(sys.prefix, "share", "xloil")

def _install_xloil():
    target_script = os.path.join(_script_dir(), "xloil_Install.ps1")
    result = subprocess.run(f"powershell {target_script}")
    if result.returncode != 0:
        raise Exception("Failed to execute install script. You may need to enable " +
                        "powershell scripts by typing 'Set-ExecutionPolicy RemoteSigned'" +
                        "at a powershell prompt with admin rights")

def _remove_xloil():
    target_script = os.path.join(_script_dir(), "xloil_Remove.ps1")
    subprocess.run(f"powershell {target_script}")

def _create_addin(args):
    if len(args) != !:
        raise Exception("'create' should have one argument, the target filename")
    target_script = os.path.join(_script_dir(), "xloil_NewAddin.ps1")
    subprocess.run(f"powershell {target_script} {args[0]}")
    
def main():
    command = sys.argv[1].lower() if len(sys.argv) > 1 else ""

    if command == 'install':
        _install_xloil()
    elif command == 'remove':
        _remove_xloil()
    elif command == 'create':
        _create_addin(sys.argv[2:])
    else:
        raise Exception("Syntax: xloil {install, remove, create}")



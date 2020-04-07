import subprocess
import sys
import os

def _script_dir():
    return os.path.join(sys.prefix, "share", "xloil")

def _install_xloil():
    install_script = os.path.join(_script_dir(), "Install_xloil.ps1")
    result = subprocess.run(f"powershell {install_script}")
    if result.returncode != 0:
        raise Exception("Failed to execute install script. You may need to enable " +
                        "powershell scripts by typing 'Set-ExecutionPolicy RemoteSigned'" +
                        "at a powershell prompt with admin rights")

def _remove_xloil():
    install_script = os.path.join(_script_dir(), "Remove_xloil.ps1")
    subprocess.run(f"powershell {install_script}")

def main():
    
    command = sys.argv[1].lower() if len(sys.argv) > 1 else ""

    if command == 'install':
        _install_xloil()
    elif command == 'remove':
        _remove_xloil()
    else:
        raise Exception("Syntax: xloil {install, remove}")



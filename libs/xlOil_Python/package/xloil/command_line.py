import sys
import os
from pathlib import Path
import re
import winreg as reg
from contextlib import suppress
import binascii
import shutil as sh

_ADDIN_NAME    = "xlOil.xll"
_INIFILE_NAME  = "xlOil.ini"
_APP_DATA_PATH = Path(os.getenv('APPDATA'))
_XL_START_PATH = _APP_DATA_PATH / "Microsoft" / "Excel" / "XLSTART"
_XLOIL_BIN_DIR = Path(sys.prefix) / "share" / "xloil"


def _excel_regpath(version):
    return rf"Software\Microsoft\Office\{version}\Excel"

def _find_excel_version():
    # Will fetch a value like "Excel.Application.16"
    # Note for Office 2019, it still gives version 16...hmm
    verstr = reg.QueryValue(reg.HKEY_CLASSES_ROOT, r"Excel.Application\CurVer")
    return int(verstr.replace('Excel.Application.',''))


def _check_VBA_access(version):

    user_access = 1
    lm_access = 1

    with suppress(OSError): 
        user_key = reg.OpenKey(reg.HKEY_CURRENT_USER, _excel_regpath(version) + r"\Security")
        user_access, = reg.QueryValueEx(user_key, "AccessVBOM")
    with suppress(OSError):
        lm_key =  reg.OpenKey(reg.HKEY_LOCAL_MACHINE, _excel_regpath(version) + r"\Security")
        lm_access, = reg.QueryValueEx(lm_key, "AccessVBOM")

    if user_access == 0 or lm_access == 0:
        print("To ensure xlOil local functions work, allow access to the VBA Object Model in\n" +
			"Excel > File > Options > Trust Center > Trust Center Settings > Macro Settings\n")


def _remove_from_resiliancy(filename, version):

    # Source https://stackoverflow.com/questions/751048/how-to-programatically-re-enable-documents-in-the-ms-office-list-of-disabled-fil

    #Converts the File Name string to UTF16 Hex
    filename_hex = binascii.hexlify(filename.encode('utf-16'))

    # If we can't find the key or exit the for loop, suppress the error
    with suppress(OSError): 
        regkey = reg.OpenKey(reg.HKEY_CURRENT_USER, _excel_regpath(version) + "\Resiliency\DisabledItems")

        for i in range(1024):
            name, value, = reg.EnumValue(regkey, i)
            value_hex = binascii.hexlify(value.encode('utf-16'))
            if filename_hex in value:
                reg.DeleteValue(regkey, name)


def _remove_addin(version):

    # If we can't find the key or exit the for loop, suppress the error
    with suppress(OSError): 
        regkey = reg.OpenKey(reg.HKEY_CURRENT_USER, _excel_regpath(version) + "\Add-in Manager")

        addin_path = _XL_START_PATH / _ADDIN_NAME
        # Cycles through all the properties and delete if it contains the file name.
        for i in range(1024):
            name, value, = reg.EnumValue(regkey, i)
            if addin_path in value:
                reg.DeleteValue(regkey, name)


def _toml_lit_string(s):
    # TOML literal strings have a lot of quotes and escapes, this function does the encoding
    return "'''" + s.replace('\\','\\\\') + "'''"


def _write_python_path_to_ini(ini_txt):

    python_path = os.path.join(sys.prefix, "Lib") + ";" + os.path.join(sys.prefix, "DLLs") 
    python_ver = f'{sys.version_info.major}.{sys.version_info.minor}'
    
    fails = 0

    def do_replace(pat, repl):
        nonlocal ini_txt, fails
        ini_txt, count = re.subn(pat, repl, ini_txt, flags=re.M)
        if count != 1:
            print(f"Failed to match pattern {pat}")
            fails += 1
    
    # Set PYTHONPATH
    do_replace(r'^(\s*PYTHONPATH\s*=).*',       r'\g<1>' + _toml_lit_string(python_path))
    # Set xlOil_PythonRoot
    do_replace(r'^(\s*xlOil_PythonRoot\s*=).*', r'\g<1>' + _toml_lit_string(sys.prefix))
    # Set xlOilPythonVersion
    do_replace(r'^(\s*PythonVersion\s*=).*',    rf'\g<1>"{python_ver}"')
    # Set XLOIL_PATH
    do_replace(r'^(\s*XLOIL_PATH\s*=).*',       r'\g<1>' + _toml_lit_string(str(_XLOIL_BIN_DIR)))

    return ini_txt, fails == 0
    
   
def _install_xloil():

    ini_path = _APP_DATA_PATH / "xlOil" / _INIFILE_NAME

    excel_version = _find_excel_version()

    # Just in case we got put in Excel's naughty corner for misbehaving addins
    _remove_from_resiliancy(_ADDIN_NAME, excel_version)

    # Check access to the VBA Object model (for local functions)
    _check_VBA_access(excel_version)

    # Ensure XLSTART dir really exists
    with suppress(FileExistsError):
        os.mkdir(_XL_START_PATH)

    # Copy the XLL
    sh.copy(_XLOIL_BIN_DIR / _ADDIN_NAME, _XL_START_PATH)
    print("Installed ", _XL_START_PATH / _ADDIN_NAME)
    
    # Copy the ini file to APPDATA, avoiding overwriting any existing ini
    if ini_path.exists():
        print("Found existing settings file at \n", ini_path)
    else:
        with suppress(FileExistsError):
            ini_path.parent.mkdir()
        sh.copy(_XLOIL_BIN_DIR / _INIFILE_NAME, ini_path)

    # Edit the xloil.ini file. To preserve comments and whitespace it's easier to just use
    # regex replace rather than read the file as structured TOML
    ini_txt = ini_path.read_text(encoding='utf-8')
    ini_txt, success = _write_python_path_to_ini(ini_txt)
    
    # Check if any of the counts is not 1, i.e. the expression matched zero or multiple times
    if not success:
        print(f'WARNING: Failed to set python paths in {ini_path}. You may have to do this manually.')
    else:
        ini_path.write_text(ini_txt, encoding='utf-8')
        print(f'Edited {ini_path} to point to {sys.prefix} python distribution.')

def _remove_xloil():

    excel_version = _find_excel_version()
    
    # Ensure no xlOil addins are in the registry
    _remove_addin(excel_version)
    
    os.remove(_XL_START_PATH / _ADDIN_NAME)


def _create_addin(args):
    if len(args) != 1:
        raise Exception("'create' should have one argument, the target filename")

    filename = args[0]
    basename = Path(os.path.splitext(filename)[0])

    xll_path = basename.with_suffix(".xll")
    ini_path = basename.with_suffix(".ini")

    sh.copy(_XLOIL_BIN_DIR / _ADDIN_NAME, xll_path)
    sh.copy(_XLOIL_BIN_DIR / "NewAddin.ini",  ini_path)
    
    print("New addin created at: ", xll_path)

    # Edit ini file
    ini_txt = ini_path.read_text(encoding='utf-8')
    
    # Assume we want the xlOil_Python plugin as we're running a python script
    ini_txt, count = re.subn(r'^(\s*Plugins\s*=).*', r'\g<1>["xlOil_Python"]', ini_txt, flags=re.M)
    
    # Assume we want the python paths set to the distribution running this script
    ini_txt, success = _write_python_path_to_ini(ini_txt)
    
    ini_path.write_text(ini_txt)

    print("xlOil_Python plugin enabled using python installed at: ", sys.prefix)


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


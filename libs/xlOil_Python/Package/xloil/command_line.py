import sys
import os
from pathlib import Path
import re
import winreg as reg
from contextlib import suppress
import binascii
import shutil as sh
from ._paths import *
from argparse import ArgumentParser

_XL_START_PATH = Path(os.getenv('APPDATA')) / "Microsoft" / "Excel" / "XLSTART"
_XLL_INSTALL_PATH = _XL_START_PATH / ADDIN_NAME

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


def _get_xloil_bin_dir():

    # We look in the a possibly overriden (by an env var) bin dir, the normal install 
    # path and the current directory
    for bin_path in [Path(XLOIL_BIN_DIR), Path(XLOIL_INSTALL_DIR), Path(".")]:
        if (bin_path / ADDIN_NAME).exists():
            return bin_path

    raise Exception(f"{ADDIN_NAME} not found")

def _remove_from_resiliancy(filename, version):

    # Source https://stackoverflow.com/questions/751048/

    #Converts the File Name string to UTF16 Hex
    filename_hex = binascii.hexlify(filename.encode('utf-16'))

    # If we can't find the key or exit the for loop, suppress the error
    with suppress(OSError): 
        regkey = reg.OpenKey(reg.HKEY_CURRENT_USER, _excel_regpath(version) + r"\Resiliency\DisabledItems")

        for i in range(1024):
            name, value, = reg.EnumValue(regkey, i)
            value_hex = binascii.hexlify(value.encode('utf-16'))
            if filename_hex in value:
                reg.DeleteValue(regkey, name)


def _remove_addin(excel_version, addin_path):

    # If we can't find the key or exit the for loop, suppress the error
    with suppress(OSError): 
        regkey = reg.OpenKey(reg.HKEY_CURRENT_USER, _excel_regpath(excel_version) + r"\Add-in Manager")

        # Cycles through all the properties and delete if it contains the file name.
        for i in range(1024):
            name, value, = reg.EnumValue(regkey, i)
            if addin_path in value:
                reg.DeleteValue(regkey, name)


def _toml_lit_string(s:str):
    # TOML literal strings have a lot of quotes and escapes, this function does the encoding
    return "'''" + s.replace('\\','\\\\') + "'''"

def _get_python_paths():
    """
        Returns the paths to be set in the xlOil.ini file (the appropriate stubs
        must already exist).
    """
    return { 
        'PYTHONEXECUTABLE': sys.executable
    }

def _write_python_path_to_ini(ini_txt, bin_dir:str, comment_reg_keys:bool, replace_paths=None):

    if replace_paths is None:
        replace_paths = _get_python_paths()

    fails = 0

    def replace(pat, repl):
        nonlocal ini_txt
        ini_txt = re.sub(pat, repl, ini_txt, count=1, flags=re.M)
            
    def check_replace(pat, repl):
        nonlocal fails, ini_txt
        if re.search(pat, ini_txt, flags=re.M) is None:
            print(f"Failed to match pattern {pat}")
            fails += 1
        else:
            replace(pat, repl)
    
    for var, value in replace_paths.items():
        check_replace(r'^(\s*' + var + r'\s*=).*', r'\g<1>' + _toml_lit_string(value))
        
    # Set XLOIL_PATH
    check_replace(r'^(\s*XLOIL_PATH\s*=).*', r'\g<1>' + _toml_lit_string(str(bin_dir)))
    
    # Comment out the now usused code to get the python paths from the registry
    # Don't error if this fails as it's not critical
    if comment_reg_keys:
        
        for key in ["xlOil_PythonRegKey"]:
            replace(rf'^(\s*{key}\s*=.*)', r'#\g<1>')

    return ini_txt, fails == 0
    
   
def install_xloil(ini_template:str=None, 
                  addin_name=None,
                  replace_ini=False, 
                  replace_paths=None):

    ini_path = Path(APP_DATA_DIR) / INIFILE_NAME

    if addin_name is None:
        addin_name = ADDIN_NAME
        
    excel_version = _find_excel_version()

    # Just in case we got put in Excel's naughty corner for misbehaving addins
    _remove_from_resiliancy(ADDIN_NAME, excel_version)

    # Check access to the VBA Object model (for local functions)
    _check_VBA_access(excel_version)

    # Ensure XLSTART dir really exists
    with suppress(FileExistsError):
        os.mkdir(_XL_START_PATH)

    bin_dir = _get_xloil_bin_dir()

    # Copy the XLL
    sh.copy(bin_dir / ADDIN_NAME, _XLL_INSTALL_PATH)
    print("Installed ", _XLL_INSTALL_PATH)
    
    # Copy the ini file to APPDATA, avoiding overwriting any existing ini
    if not replace_ini and ini_path.exists():
        print("Found existing settings file at \n", ini_path)
    else:
        with suppress(FileExistsError):
            ini_path.parent.mkdir()
        sh.copy(ini_template or bin_dir / INIFILE_NAME, ini_path)

    # Edit the xloil.ini file. To preserve comments and whitespace it's easier to just use
    # regex replace rather than read the file as structured TOML
    ini_txt = ini_path.read_text(encoding='utf-8')

    ini_txt, success = _write_python_path_to_ini(ini_txt, bin_dir, 
                                                 comment_reg_keys=True, 
                                                 replace_paths=replace_paths)

    # Check if any of the counts is not 1, i.e. the expression matched zero or multiple times
    if not success:
        print(f'WARNING: Failed to set python paths in {ini_path}. You may have to do this manually.')
    else:
        ini_path.write_text(ini_txt, encoding='utf-8')
        print(f'Edited {ini_path} to point to {sys.prefix} python distribution.')

def _remove_xloil():

    excel_version = _find_excel_version()
    
    # Ensure no xlOil addins are in the registry
    _remove_addin(excel_version, _XLL_INSTALL_PATH)
    try:
        os.remove(_XLL_INSTALL_PATH)
    except FileNotFoundError:
        ...

def _clean_xloil():

    _remove_xloil()

    ini_path = Path(APP_DATA_DIR) / INIFILE_NAME
    try:
        os.remove(os.path.join(APP_DATA_DIR, INIFILE_NAME))
        os.remove(os.path.join(APP_DATA_DIR, "xlOil.log"))
    except FileNotFoundError:
        ...
    import subprocess
    import sys

    subprocess.Popen(f"{sys.executable} -m pip uninstall --yes xloil", shell=True)


def _create_addin(filename):

    basename = Path(os.path.splitext(filename)[0])

    xll_path = basename.with_suffix(".xll")
    ini_path = basename.with_suffix(".ini")

    bin_dir = _get_xloil_bin_dir()
    print("xlOil binaries found at:", str(bin_dir))

    sh.copy(bin_dir / ADDIN_NAME,    xll_path)
    sh.copy(bin_dir / INIFILE_NAME,  ini_path)
    
    print("New addin created at:", xll_path)

    # Edit ini file
    ini_txt = ini_path.read_text(encoding='utf-8')
    
    # Assume we want the xlOil_Python plugin as we're running a python script
    ini_txt, count = re.subn(r'^(\s*Plugins\s*=).*', r'\g<1>["xlOil_Python"]', ini_txt, flags=re.M)

    # Assume we want the python paths set to the distribution running this script
    ini_txt, success = _write_python_path_to_ini(ini_txt, bin_dir, 
                                                 comment_reg_keys=True)
    
    ini_path.write_text(ini_txt)

    print("xlOil_Python plugin enabled using python installed at: ", sys.prefix)
 
    
def _package_pyinstaller(ini_template: str, 
                         makespec: bool = False, 
                         extra_args: str = None):

    if ini_template is None:
        ini_template = Path(APP_DATA_DIR) / INIFILE_NAME
    else:
        ini_template = Path(ini_template)
        
    python3_dll = os.path.join(sys.base_exec_prefix, "python3.dll")

    # Use the filename stems as when the script is run, these files will be
    # available in its current directory. 
    ini_filename = ini_template.name
    addin_filename = "xloil.xll"
    
    entry_point = "install_main.py"
    
    install_stub = \
        'import os\n' \
        'import sys\n' \
        'from xloil.command_line import install_xloil\n' \
        f'install_xloil(ini_template="{ini_filename}", addin_name="{addin_filename}",\n' \
        '   replace_paths={"PYTHONEXECUTABLE": os.path.join(sys.prefix, "python.exe")})\n'
    
    Path(entry_point).write_text(install_stub)

    args = [
        '',
        entry_point,
        '--onedir',
        '--debug=noarchive',
        '--nowindow',
        f'--add-data={ini_template}:.',
        f'--add-data={python3_dll}:.',
        f'--add-data={sys.executable}:.',
    ]

    if extra_args is not None:
        args += extra_args
       
    print(args)
    
    try:
        if makespec:
            from PyInstaller.utils.cliutils.makespec import run
        else:
            from PyInstaller.__main__ import run
    except ImportError:
        raise ImportError("To run `xloil package` you need the pyinstaller package installed")

    sys.argv = args
    run()


def main():
    parser = ArgumentParser(prog='xlOil', 
                            description='Excel/Python integration')
    
    subparsers = parser.add_subparsers(help='Commands to control local xlOil distribution')

    p = subparsers.add_parser('install', 
                              help='Installs xloil by copying the Excel addin to the XLSTART directory '
                                   'and sets the correct paths in xloil.ini')
    p.add_argument("ini_template", nargs='?', type=str, default=None,
                   help='Path to xloil.ini file to modify and install, if not present the default file is used')
    p.set_defaults(func=install_xloil)
    
    p = subparsers.add_parser('create', 
                              help='Creates an XLL addin to simplify distribution of Excel funcs to other xlOil users')
    p.add_argument("filename", type=str, 
                   help='Name of the XLL file to create')
    p.set_defaults(func=_create_addin)

    p = subparsers.add_parser('package', 
                              help='Uses PyInstaller to create a packaged xloil installer. Any arguments '
                                   'which xlOil does not parse will be passed directly to PyInstaller. '
                                   'Invoke this function from a minimal python distribution, otherwise the '
                                   'resulting package may be large. See docs for more details.')
    p.add_argument("ini_template", type=str, 
                   help='The xlOil ini file to package and distribute')
    p.add_argument("--makespec", action='store_true', 
                   help='If present, creates the PyInstaller spec file then exits. This allows more precise '
                        'tweaking of the PyInstaller config. See PyInstaller docs on pyi-makespec')
    p.set_defaults(func=_package_pyinstaller)

    p = subparsers.add_parser('clean', 
                              help='Uninstalls xlOil and removes its package using pip')
    p.set_defaults(func=_clean_xloil)

    p = subparsers.add_parser('remove', 
                              help='Removes the xlOil addin from Excel. Leaves xlOil.ini in place')
    p.set_defaults(func=_remove_xloil)
    
    p = subparsers.add_parser('uninstall', 
                              help='Equivalent to the remove command')
    p.set_defaults(func=_remove_xloil)
    
    args, more_args = parser.parse_known_args()
    
    func_args = vars(args)
    func = func_args.pop("func")
    if len(more_args) > 0:
        func(**vars(args), extra_args=more_args)
    else:
        func(**vars(args))
        

if __name__ == '__main__':
    main()
import os
from pathlib import Path
import sys
import subprocess
import distutils.dir_util
import distutils.file_util
from glob import glob
import shutil as sh
from argparse import ArgumentParser

def merge_dict(x, y):
    return {**x, **y}

parser = ArgumentParser()
parser.add_argument("--post-ver")
parser.add_argument("--no-build", action='store_true')
cmd_args,_ = parser.parse_known_args()

tools_dir = Path(os.path.realpath(__file__)).parent
soln_dir = tools_dir.parent
doc_dir = soln_dir / "docs"
build_dir = soln_dir / "build"
staging_dir = build_dir / "staging"
plugin_dir = soln_dir / "libs"
include_dir = soln_dir / "include"

architectures = ["x64", "Win32"]

python_versions = ["3.6", "3.7", "3.8", "3.9", "3.10"]
python_package_dir = staging_dir / "pypackage"

build_files = {}
build_files['x64'] = {
    'Core' : ["xlOil.xll", "xlOil.dll", "xlOil.lib"],
    'xlOil_Python': ["xlOil_Python.dll"] + [f"xlOil_Python{ver.replace('.','')}.pyd" for ver in python_versions],
    'xlOil_SQL': ["xlOil_SQL.dll"],
    'xlOil_Utils': ["xlOil_Utils.dll"] 
}
build_files['Win32'] = build_files['x64'].copy()

lib_files = [
     { 
        'from': '.',
        'files': ['Version.txt'],
        'to': '.'
    },
    { 
        'from': 'tools',
        'files': ['xlOil_Install.ps1', 'xlOil_Remove.ps1', 'xlOil_NewAddin.ps1'],
        'to': architectures
    },
    { 
        'from': 'config',
        'files': ['xloil.ini'],
        'to': architectures
    },
    {
        'from': 'libs/xlOil_Python',
        'files': ['package', soln_dir / "LICENSE"],
        'to': 'pypackage'
    },
    {
        'from': '.',
        'files': ['README.md'],
        'to': 'pypackage'
    },
    {
        'from': "include",
        'files': '.',
        'to': 'include'
    },
]


def copy_tree(src, dst):
    distutils.dir_util.copy_tree(str(src), str(dst))

def copy_file(src, dst):
    distutils.file_util.copy_file(str(src), str(dst))

def latest_file(dir):
    list_of_files = glob(f'{dir}/*')
    return max(list_of_files, key=os.path.getctime)


print("Soln dir: ", str(soln_dir))

# Write the version file
subprocess.run(f"powershell ./tools/WriteVersion.ps1", cwd=soln_dir, check=True)

if not 'no_build' in cmd_args or cmd_args.no_build is False:
    # Build the library
    for arch in architectures:
        subprocess.run(f"BuildRelease.cmd {arch}", cwd=tools_dir, check=True)

    # Write the combined include file
    subprocess.run(f"powershell ./WriteInclude.ps1 {include_dir / 'xloil'} {staging_dir / 'include' / 'xloil'}", 
                   cwd=tools_dir, check=True)

# Build the docs
# TODO: check=True should throw if the process exit code is != 0. Doesn't work.
subprocess.run(f"cmd /C make.bat doxygen", cwd=doc_dir, check=True)
subprocess.run(f"cmd /C make.bat -bin x64\Release html", cwd=doc_dir, check=True)

#
# Start of file copying
#

# Clean any previous python package
try: sh.rmtree(python_package_dir)
except (FileNotFoundError, OSError): pass


for arch in architectures:
    for files in build_files[arch].values():
        for file in files:
            try: os.makedirs(staging_dir / arch)
            except FileExistsError: pass
            copy_file(build_dir / arch / "Release" / file, staging_dir / arch)


for job in lib_files:
    source = soln_dir / job['from']
    print(source)
    targets = job['to']
    if not isinstance(job['to'], list):
        targets = [targets]
    for target in targets:
        for f in job['files']:
            print(" ", f)
            target_path = staging_dir / target
            try: os.makedirs(target_path)
            except FileExistsError: pass
            if os.path.isabs(f):
                copy_file(f, target_path)
            elif os.path.isdir(source / f):
                copy_tree(source / f, target_path)
            else:
                copy_file(source/ f, target_path)

copy_tree(doc_dir / "source" / "_build" / "html", staging_dir / "docs")

#
# Create distributable archives
#
import tarfile

xloil_version =  (soln_dir / 'Version.txt').read_text().replace('\n','')

for arch in architectures:
    with tarfile.open(staging_dir / f"xlOil-{xloil_version}-{arch}-bin.tar.bz2", "w:bz2") as tar:
        tar.add(staging_dir / arch, arcname=arch)

with tarfile.open(staging_dir / f"xlOil-{xloil_version}-docs.tar.bz2", "w:bz2") as tar:
    tar.add(staging_dir / "docs", arcname='docs')

with tarfile.open(staging_dir / f"xlOil-{xloil_version}-include.tar.bz2", "w:bz2") as tar:
    tar.add(staging_dir / "include", arcname='include')

#
# Build python wheels
#
for arch in architectures:
    platform_tags = { 'Win32': 'win32', 'x64': 'win_amd64'}
    plat_name = platform_tags[arch]

    our_pytag = f'cp{sys.version_info.major}{sys.version_info.minor}'

    pypi_version = xloil_version
    if 'post_ver' in cmd_args and cmd_args.post_ver is not None:
        pypi_version += f'.post{cmd_args.post_ver}'
       
    for pyver in python_versions:
        # It's important to run the setup using the targeted python version
        # If you get errors building win32 on an x64 version of python,
        # just comment out the assert in get_tag() in bdist_wheel.py.
        # Guido probably wouldn't approve but it seems to work.
        cmd = f"py -{pyver} setup.py bdist_wheel --arch {arch} --pyver {pyver} --version {pypi_version} --plat-name {plat_name}"
        print(f"Running: {cmd}.")
        subprocess.run(cmd, cwd=f"{python_package_dir}")

#
# Next steps
#
print(
     '\n'
     '\nTo test the python package:'
    f'\n  > pip install {str(python_package_dir)}\\dist\\<wheel file>'
     '\n  > xloil install'
    r'\n  > python ..\libs\xlOil_Python\Package\test_PythonAutomation.py'
    r'\n  > python ..\libs\xlOil_Python\Package\test_SpreadsheetRunner.py'
    '\n'
    '\nTo upload the python package to PyPI:'
    f'\n  > cd {str(python_package_dir)}'
     '\n  > twine upload --repository-url https://test.pypi.org/legacy/ dist/*'
    '\nor'
     '\n  > twine upload dist/*'
    )

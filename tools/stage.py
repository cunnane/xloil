import os
from pathlib import Path
import sys
import subprocess
import distutils.dir_util
import distutils.file_util
from glob import glob
import shutil as sh

def merge_dict(x, y):
    return {**x, **y}

tools_dir = Path(os.path.realpath(__file__)).parent
soln_dir = tools_dir.parent
doc_dir = soln_dir / "docs"
build_dir = soln_dir / "build"
staging_dir = build_dir / "staging"
plugin_dir = soln_dir / "libs"
include_dir = soln_dir / "include"

architectures = ["x64"]

python_versions = ["3.6", "3.7"]
python_package_dir = staging_dir / "pypackage"

build_files = {
    'Core' : ["xlOil.xll", "xlOil.dll", "xloil.ini"],
    'xlOil_Python': ["xlOil_Python36.dll", "xlOil_Python37.dll", "xlOil_Python.dll"],
    'xlOil_SQL': ["xlOil_SQL.dll"],
    'xlOil_Utils': ["xlOil_Utils.dll"] 
}

lib_files = [
    { 
        'from': 'tools',
        'files': ['xlOil_Install.ps1', 'xlOil_Remove.ps1', 'xlOil_NewAddin.ps1'],
        'to': architectures
    },
    { 
        'from': 'src',
        'files': ['NewAddin.ini'],
        'to': architectures
    },
    {
        'from': 'libs/xlOil_Python',
        'files': ['package', soln_dir / "LICENSE"],
        'to': 'pypackage'
    },
    {
        'from': 'libs/xlOil_Python',
        'files': ['package/xloil/xloil.py'],
        'to': architectures
    }
]


def copy_tree(src, dst):
    distutils.dir_util.copy_tree(str(src), str(dst))

def copy_file(src, dst):
    distutils.file_util.copy_file(str(src), str(dst))

def latest_file(dir):
    list_of_files = glob(f'{dir}/*')
    return max(list_of_files, key=os.path.getctime)


print("Soln dir: ", str(soln_dir))

# Build the library
subprocess.run(f"BuildRelease.cmd", cwd=tools_dir)

# Write the combined include file
subprocess.run(f"powershell ./WriteInclude.ps1 {include_dir} {include_dir}", cwd=tools_dir)

# Build the docs
subprocess.run(f"cmd /C make.bat html", cwd=doc_dir)

#
# Start of file copying
#

for files in build_files.values():
    for file in files:
        for arch in architectures:
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
            if os.path.isabs(f):
                copy_file(f, target_path)
            elif os.path.isdir(source / f):
                copy_tree(source / f, target_path)
            else:
                copy_file(source/ f, target_path)

copy_tree(doc_dir / "build" / "html", staging_dir / "docs")

#
# Build python wheels
#
try:
    sh.rmtree(python_package_dir / "dist")
except:
    pass

for arch in architectures:
    for pyver in python_versions:
        subprocess.run(f"python setup.py bdist_wheel --arch {arch} --pyver {pyver}", cwd=f"{python_package_dir}")# --python-tag py2 --plat-name x86")
        wheel = Path(latest_file(python_package_dir / "dist"))
        verXY = pyver.replace('.','')
        correct_name = wheel.name.replace("cp37", f'cp{verXY}')
        print("Renaming:", wheel, correct_name)
        os.rename(wheel, python_package_dir / "dist" / correct_name)

# pip wheel -w dist
# twine upload dist/*
#twine upload --repository-url https://test.pypi.org/legacy/ dist/*



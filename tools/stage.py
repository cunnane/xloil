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
        'from': '.',
        'files': ['README.md'],
        'to': 'pypackage'
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

# Write the version file
subprocess.run(f"powershell ./tools/WriteVersion.ps1", cwd=soln_dir)

# Build the docs
subprocess.run(f"cmd /C make.bat html", cwd=doc_dir)

#
# Start of file copying
#

# Clean any previous python package
try: sh.rmtree(python_package_dir)
except (FileNotFound): pass


for files in build_files.values():
    for file in files:
        for arch in architectures:
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
    with tarfile.open(staging_dir / f"xlOil-{xloil_version}-{arch}.tar.bz2", "w:bz2") as tar:
        tar.add(staging_dir / arch, arcname=arch)

with tarfile.open(staging_dir / f"xlOil-{xloil_version}-docs.tar.bz2", "w:bz2") as tar:
    tar.add(staging_dir / "docs", arcname='docs')

#
# Build python wheels
#
for arch in architectures:
    for pyver in python_versions:
        subprocess.run(f"python setup.py bdist_wheel --arch {arch} --pyver {pyver}", cwd=f"{python_package_dir}")# --python-tag py2 --plat-name x86")
        wheel = Path(latest_file(python_package_dir / "dist"))
        verXY = pyver.replace('.','')
        correct_name = wheel.name.replace("cp37", f'cp{verXY}')
        print("Renaming:", wheel, correct_name)
        os.rename(wheel, python_package_dir / "dist" / correct_name)

#
# Next steps
#
print('\n\nTo upload the python package to PyPI:')
print(f'cd {str(python_package_dir)}')
print('twine upload --repository-url https://test.pypi.org/legacy/ dist/*')
print('  or ')
print('twine upload dist/*')

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

architectures = ["x64", "Win32"]

python_versions = ["3.6", "3.7", "3.8"]
python_package_dir = staging_dir / "pypackage"

build_files = {}
build_files['x64'] = {
    'Core' : ["xlOil.xll", "xlOil.dll", "xloil.ini"],
    'xlOil_Python': ["xlOil_Python36.dll", "xlOil_Python37.dll", "xlOil_Python38.dll", "xlOil_Python.dll"],
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
for arch in architectures:
    subprocess.run(f"BuildRelease.cmd {arch}", cwd=tools_dir)

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
    with tarfile.open(staging_dir / f"xlOil-{xloil_version}-{arch}.tar.bz2", "w:bz2") as tar:
        tar.add(staging_dir / arch, arcname=arch)

with tarfile.open(staging_dir / f"xlOil-{xloil_version}-docs.tar.bz2", "w:bz2") as tar:
    tar.add(staging_dir / "docs", arcname='docs')

#
# Build python wheels
#
for arch in architectures:
    platform_tags = { 'Win32': 'win32_foo', 'x64': 'win_amd64_foo'}
    plat_name = platform_tags[arch]
    our_pytag = f'cp{sys.version_info.major}{sys.version_info.minor}'
    for pyver in python_versions:
        #
        # We need the foo suffix because setup ignores the --python-tag specification so we have
        # to manually rename the files. Glorious automation.
        #
        
        cmd = f"python setup.py bdist_wheel --arch {arch} --pyver {pyver} --plat-name {plat_name}"
        print(f"Running: {cmd}.")
        subprocess.run(cmd, cwd=f"{python_package_dir}")

        wheel = Path(latest_file(python_package_dir / "dist"))
        verXY = pyver.replace('.','')
        ### TODO: the cp37 depends on the current py version in use
        correct_name = wheel.name.replace(our_pytag, f'cp{verXY}').replace('_foo','')
        os.rename(wheel, python_package_dir / "dist" / correct_name)

#
# Next steps
#
print('\n\nTo upload the python package to PyPI:')
print(f'cd {str(python_package_dir)}')
print('twine upload --repository-url https://test.pypi.org/legacy/ dist/*')
print('  or ')
print('twine upload dist/*')

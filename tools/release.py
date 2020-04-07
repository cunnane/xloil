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



soln_dir = Path(os.path.realpath(__file__)).parent.parent
print("Soln: ", str(soln_dir))
doc_dir = soln_dir / "docs"
build_dir = soln_dir / "build"
staging_dir = build_dir / "staging"
plugin_dir = soln_dir / "libs"

architectures = ["x64"]

build_files = {
    'Core' : ["xlOil.xll", "xlOil.dll", "xloil.ini"],
    'xlOil_Python': ["xlOil_Python36.dll", "xlOil_Python37.dll", "xlOil_Python36.ini", "xlOil_Python37.ini"],
    'xlOil_SQL': ["xlOil_SQL.dll"],
    'xlOil_Utils': ["xlOil_Utils.dll"] 
}

lib_files = {
    'tools' : {arch : ['Install_xlOil.ps1', 'Remove_xlOil.ps1'] for arch in architectures},
    'libs/xlOil_Python' : merge_dict(
    {
        'pypackage' : ['package', soln_dir / "LICENSE"]
    },
    {
        arch : ['package/xloil/xloil.py'] for arch in architectures
    })
}

python_versions = ["3.6", "3.7"]
python_package_dir = staging_dir / "pypackage"


def copy_tree(src, dst):
    distutils.dir_util.copy_tree(str(src), str(dst))

def copy_file(src, dst):
    distutils.file_util.copy_file(str(src), str(dst))

def latest_file(dir):
    list_of_files = glob(f'{dir}/*')
    return max(list_of_files, key=os.path.getctime)
    
for files in build_files.values():
    for file in files:
        for arch in architectures:
            copy_file(build_dir / arch / "Release" / file, staging_dir / arch)

for plugin, targets in lib_files.items():
    print(plugin)
    for target_dir, sources in targets.items():
        for f in sources:
            print(f)
            target_path = staging_dir / target_dir
            if os.path.isabs(f):
                copy_file(f, target_path)
            elif os.path.isdir(soln_dir / plugin / f):
                copy_tree(soln_dir / plugin / f, target_path)
            else:
                copy_file(soln_dir / plugin / f, target_path)

sh.rmtree(python_package_dir / "dist")
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

#import mkdocs

#sh.copy(doc_dir, staging_dir)

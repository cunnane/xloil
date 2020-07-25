from setuptools import setup, Distribution
import sys
from pathlib import Path
from argparse import ArgumentParser
import os
import shutil as sh
import re

#
# Process our cmd line args
#
parser = ArgumentParser()
parser.add_argument("--arch")
parser.add_argument("--pyver")
args, unknown = parser.parse_known_args()

if 'arch' not in args:
    raise Exception("No architecture specified")

if 'pyver' not in args:
    raise Exception("No python version specified")

target_py_ver = args.pyver

# Pass the un-parsed args to setuptools
sys.argv = [sys.argv[0]] + unknown


#
# Define directoies
#
 
staging_dir = Path('..')
bin_dir = staging_dir / args.arch


#
# Specify data files to copy
#
data_files = [str(bin_dir / f) for f in [
    'xlOil.xll', 
    'xlOil.ini', 
    'xlOil.dll',
    'xlOil_Python.dll', 
    'xlOil_Utils.dll', 
    'xlOil_SQL.dll', 
    'xlOil_Install.ps1', 
    'xlOil_NewAddin.ps1',
    'xlOil_Remove.ps1']]

verXY = target_py_ver.replace('.','')
data_files += [str(bin_dir / f'xlOil_Python{verXY}.dll')]

#
# Special treatment for ini file
# 
try: os.makedirs(verXY)
except FileExistsError: pass

ini_path = Path(verXY) / 'xlOil.ini'
sh.copyfile(bin_dir / 'xlOil.ini', ini_path)

# Fix up the version number in the ini file
ini_text = ini_path.read_text()
ini_text = re.sub(r'(xlOilPythonVersion=\")[0-9.]+(\")', f"\\g<1>{target_py_ver}\\g<2>", ini_text)
ini_path.write_text(ini_text)

data_files += [str(ini_path)]

#
# Grab the help text from README.md
# 
with open("README.md", "r") as fh:
    contents_of_readme = fh.read()

class BinaryDistribution(Distribution):
    """Distribution which always forces a binary package with platform name"""
    def has_ext_modules(self):
        return True

version = Path(staging_dir / "Version.txt").read_text()
setup(
    name="xlOil",
    version=version,
    author="Steven Cunnane",
    author_email="my-surname@gmail.com",
    description="Excel interop for Python and Jupyter",
    long_description=contents_of_readme,
    long_description_content_type="text/markdown",
    url="https://gitlab.com/stevecu/xloil",
    download_url='https://gitlab.com/stevecu/xloil/-/releases/',
    project_urls = {
      'Documentation': 'https://xloil.readthedocs.io',
    },
    license='Apache',
    
    distclass=BinaryDistribution,
    packages=['xloil'],
    data_files=[('share/xloil', data_files)],
    entry_points = {
        'console_scripts': ['xloil=xloil.command_line:main'],
    },

    # Doesn't work, but the internet says it should
    # options={'bdist_wheel':{'python_tag':'foo'}},
    
    python_requires=f'>={target_py_ver}',
    install_requires=[
        'numpy'
    ],
    
    classifiers=[
        'Development Status :: 3 - Alpha',
        "Programming Language :: Python :: 3",
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: C++',
        'Topic :: Software Development :: Libraries :: Python Modules',
        'Topic :: Office/Business :: Financial :: Spreadsheet',
        'Framework :: Jupyter',
        'Operating System :: Microsoft :: Windows'
    ]
)
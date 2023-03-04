from setuptools import setup, Distribution, find_packages
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
parser.add_argument("--version")
args, unknown = parser.parse_known_args()

if 'arch' not in args:
    raise Exception("No architecture specified")

if 'pyver' not in args:
    raise Exception("No python version specified")

if 'version' not in args:
    raise Exception("No xloil version specified")

target_py_ver = args.pyver

# Pass the un-parsed args to setuptools
sys.argv = [sys.argv[0]] + unknown

#
# Define directoies
#
# We are run from ./staging/pypackage so need to go up a directory

staging_dir = staging_dir = Path(__file__).resolve().parent.parent
bin_dir = staging_dir / args.arch

#
# Specify data files to copy
#
data_files = [str(bin_dir / f) for f in [
    'xlOil.xll', 
    'xlOil.dll',
    'xlOil_Python.dll', 
    'xlOil_Utils.dll', 
    'xlOil_SQL.dll', 
    'xlOil_Install.ps1', 
    'xlOil_NewAddin.ps1',
    'xlOil_Remove.ps1']]

data_files += [str(staging_dir / 'xlOil.ini')]

py_version_XY = target_py_ver.replace('.','')
data_files += [str(bin_dir / f'xlOil_Python{py_version_XY}.pyd')]

#
# Grab the help text from README.md
# 
with open("README.md", "r") as fh:
    contents_of_readme = fh.read()

class BinaryDistribution(Distribution):
    """Distribution which always forces a binary package with platform name"""
    def has_ext_modules(self):
        return True

setup(
    name="xlOil",
    version=args.version,
    author="Steven Cunnane",
    author_email="replace-this-with-my-surname@gmail.com",
    description="Excel interop for Python and Jupyter",
    long_description=contents_of_readme,
    long_description_content_type="text/markdown",
    url="https://gitlab.com/stevecu/xloil",
    download_url='https://gitlab.com/stevecu/xloil/-/releases/',
    project_urls = {
      'Documentation': 'https://xloil.readthedocs.io/en/stable',
    },
    license='Apache',
    
    distclass=BinaryDistribution,
    packages=find_packages(),
    data_files=[('share/xloil', data_files)],
    entry_points = {
        'console_scripts': ['xloil=xloil.command_line:main'],
    },

    # Doesn't work, but the internet says it should
    # options={'bdist_wheel':{'python_tag':'foo'}},
    
    python_requires=f'>=3.6',
    install_requires=[
        'numpy>=1.18',
        'tomlkit'
    ],
    
    classifiers=[
        'Development Status :: 4 - Beta',
        "Programming Language :: Python :: 3",
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
        'Programming Language :: C++',
        'Topic :: Software Development :: Libraries :: Python Modules',
        'Topic :: Office/Business :: Financial :: Spreadsheet',
        'Topic :: Software Development :: Libraries :: Application Frameworks',
        'Framework :: Jupyter',
        'Operating System :: Microsoft :: Windows'
    ]
)
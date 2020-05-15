from setuptools import setup, Distribution
import sys
from pathlib import Path
from argparse import ArgumentParser

# Process our cmd line args

parser = ArgumentParser()
parser.add_argument("--arch")
parser.add_argument("--pyver")
args, unknown = parser.parse_known_args()

# Pass the un-parsed args to setuptools
sys.argv = [sys.argv[0]] + unknown


if 'arch' not in args:
    raise Exception("No architecture specified")

if 'pyver' not in args:
    raise Exception("No python version specified")

bin_dir = Path('..') / args.arch

target_py_ver = args.pyver

data_files = [str(bin_dir / f) for f in [
    'xlOil.xll', 
    'xlOil.ini', 
    'xlOil.dll',
    'xlOil_Python.dll', 
    'xlOil_Install.ps1', 
    'xlOil_NewAddin.ps1',
    'xlOil_Remove.ps1']]

verXY = target_py_ver.replace('.','')
data_files += [str(bin_dir / f'xlOil_Python{verXY}.dll')]


with open("README.md", "r") as fh:
    contents_of_readme = fh.read()

class BinaryDistribution(Distribution):
    """Distribution which always forces a binary package with platform name"""
    def has_ext_modules(self):
        return True

setup(
    name="xlOil",
    version="0.2",
    author="Steven",
    author_email="cunnane@gmail.com",
    description="Excel interface layer and things",
    long_description=contents_of_readme,
    long_description_content_type="text/markdown",
    url="https://gitlab.com/stevecu/xloil",
    download_url='https://gitlab.com/stevecu/xloil/-/releases/0.16-alpha',
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
        'Topic :: Software Development :: Libraries :: Python Modules'
    ]
)
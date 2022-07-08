#
# Called from xloil/docs/make.cmd
#

import os
import sys
from pathlib import Path

PACKAGE_PATH = Path(__file__).parent.resolve()
SOLUTION_PATH = PACKAGE_PATH / "../../../"

# Setup paths to import xlOil
sys.path.append(str(PACKAGE_PATH))
import xloil

# Setup paths to import pybind11_stubgen
sys.path.append(str(SOLUTION_PATH / "external"))
from pybind11_stubgen import ModuleStubsGenerator, DirectoryWalkerGuard

# Run the stub generator
out_dir = PACKAGE_PATH / 'stubs'
mod = ModuleStubsGenerator('xloil_core')
mod.parse()
with DirectoryWalkerGuard(out_dir):
    mod.write()

# Seems awfully hard to copy/delete directories in python, you'd
# think pathlib would have it covered
from distutils.dir_util import copy_tree
copy_tree(str(out_dir), str(PACKAGE_PATH / "xloil" / "stubs"))
import shutil
shutil.rmtree(str(out_dir))
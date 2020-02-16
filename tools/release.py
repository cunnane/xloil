import os
from pathlib import Path
import shutil as sh
import sys

soln_dir = Path(os.path.realpath(__file__)).parent.parent
print("Soln: ", str(soln_dir))
doc_dir = soln_dir / "docs"
build_dir = soln_dir / "build"
staging_dir = build_dir / "staging"

architectures = ["x64"]

arch_files = {
    'Core' : {
        'build': ["xlOil.xll", "xlOil.dll", "xloil.ini"]
    },
    'Python': {
        'build': ["xlOil_Python36.dll", "xlOil_Python37.dll", "xlOil_Python36.ini", "xlOil_Python37.ini"],
        'more': [soln_dir / "libs" / "xlOil_Python" / "xloil.py"]
    }
}


for source in arch_files.values():
    for file in source.get('build', []):
        for arch in architectures:
            sh.copy(build_dir / arch / "Release" / file, staging_dir / arch)
    for pth in source.get('more', []):
        sh.copy(pth, staging_dir / arch)



import mkdocs

#sh.copy(doc_dir, staging_dir)

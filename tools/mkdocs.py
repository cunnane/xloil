import os
from pathlib import Path

soln_dir = Path(os.path.realpath(__file__)).parent.parent
doc_dir = soln_dir / "docs"
py_dir = soln_dir / "src" / "xlOil_Python"
print("xloil_python", py_dir)

os.system(f"cd {py_dir} & pdoc3 --html --force --output-dir {doc_dir} xloil")
os.remove(doc_dir / "xlOil_Python.html")
os.rename(doc_dir / "xloil.html", doc_dir / "xlOil_Python.html")

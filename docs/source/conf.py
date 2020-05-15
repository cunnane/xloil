# Configuration file for the Sphinx documentation builder.
#
# This file only contains a selection of the most common options. For a full
# list see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Path setup --------------------------------------------------------------

# If extensions (or modules to document with autodoc) are in another directory,
# add these directories to sys.path here. If the directory is relative to the
# documentation root, use os.path.abspath to make it absolute, like shown here.
#
# import os
# import sys
# sys.path.insert(0, os.path.abspath('.'))


# -- Project information -----------------------------------------------------

project = 'xlOil'
copyright = '2020, Steven Cunnane'
author = 'Steven Cunnane'

# The full version, including alpha/beta/rc tags
release = '0.3'

import os
import sys
from pathlib import Path

soln_dir = Path(os.path.realpath(__file__)).parent.parent.parent
print("xlOil solution directory: ", str(soln_dir))
sys.path.append(str(soln_dir / "libs" / "xlOil_Python" / "package"))

# -- General configuration ---------------------------------------------------

# Add any Sphinx extension module names here, as strings. They can be
# extensions coming with Sphinx (named 'sphinx.ext.*') or your custom
# ones.
extensions = ["sphinx.ext.autodoc", "sphinx.ext.autosummary"]

# Add any paths that contain templates here, relative to this directory.
templates_path = ['_templates']

# List of patterns, relative to source directory, that match files and
# directories to ignore when looking for source files.
# This pattern also affects html_static_path and html_extra_path.
exclude_patterns = []


# -- Options for HTML output -------------------------------------------------

# The theme to use for HTML and HTML Help pages.  See the documentation for
# a list of builtin themes.
#
html_theme = 'bizstyle'

# Add any paths that contain custom static files (such as style sheets) here,
# relative to this directory. They are copied after the builtin static files,
# so a file named "default.css" will overwrite the builtin "default.css".
html_static_path = ['_static']

# A list of paths that contain extra files not directly related to the documentation, 
# such as robots.txt or .htaccess. Relative paths are taken as relative to the 
# configuration directory. They are copied to the output directory.
#html_extra_path = ['../build/doxygen']


autodoc_default_flags = ['members']

autosummary_generate = True

# -- Generate examples file ---------------------------------------------------

import zipfile
from zipfile import ZipFile

zipObj = ZipFile('../build/xlOilExamples.zip', 'w', compression=zipfile.ZIP_BZIP2)
 
zipObj.write(soln_dir / "tests" / "python" / "PythonTest.xlsm", "PythonTest.xlsm")
zipObj.write(soln_dir / "tests" / "python" / "PythonTest.py", "PythonTest.py")
zipObj.write(soln_dir / "tests" / "sql" / "TestSQL.xlsx", "TestSQL.xlsx")
zipObj.write(soln_dir / "tests" / "utils" / "xlOil_Utils.xlsx", "xlOil_Utils.xlsx")
 
zipObj.close()

# -- Build Doxygen Docs -----------------------------------------------------------

# These build to ../build/doxygen

import subprocess
subprocess.call('doxygen xloil.doxyfile', shell=True, cwd=soln_dir / "docs")


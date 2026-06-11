=======================
xlOil Development Guide
=======================

Getting started as developer
----------------------------

- You need Visual Studio 2022 or newer
- Run `scripts\DownloadDependencies.cmd` to fetch external dependencies from github
- You need to manually unpack `external\python.7z` if you want to build the python libs.
- To build the `xlOil-COM` library you need the ATL headers which can be installed with the Visual
  Studio installer (under C++ development). 

Debugging
=========
  
For C++ native debugging, set xlOil_Loader as the target project and point the debugger to your *Excel.exe*, typically:
  * *command*=`C:\Program Files\Microsoft Office\root\Office16\Excel.exe` 
  * *args*=`$(OutDir)\xloil.xll`
  * *environment*=`PYTHONPATH=$(SolutionDir)libs\xlOil_Python\Package;XLOIL_SETTINGS_DIR=$(OutDir)$(LocalDebuggerEnvironment)`

Mixed mode python/C++ debugging may now work, but at the time of writing, it takes a very long time to startup, then fails to hit many python breakpoints.


Release Instructions
--------------------

::

    cd <xloil_root>\tools
    python stage.py

(Optional) test python wheels with 

::

    cd <xloil_root>\build\staging\pypackage
    pip install dist/xlOil-0.3-cp37-cp37m-win_amd64.whl
    pip uninstall dist/xlOil-0.3-cp37-cp37m-win_amd64.whl

Use twine to upload to PyPI (note you need to ensure twine has the right login
keys/secrets):

::

    cd <xloil_root>\build\staging\pypackage

    # (Optional test)
    twine upload --repository-url https://test.pypi.org/legacy/ dist/*

    # The real thing
    twine upload dist/*

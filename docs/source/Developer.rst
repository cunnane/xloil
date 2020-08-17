=======================
xlOil Development Guide
=======================

Getting started as developer
----------------------------

- You need Visual Studio 2017 or newer
- All xlOil_Core dependencies are already in the `external` folder. 
  Some of them are compressed, so unpack them.
- To build xlOil_Python you need to set the correct paths in 
  `PySettings.props`, `PySettings36.props` and `PySettings37.props` to point to your python distribution.
- For debugging, set xlOil_Loader as the target project, with 
  command=`<Path-to-Excel.exe>` args=`$(OutDir)\xloil.xll`


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

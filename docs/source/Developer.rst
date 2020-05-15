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



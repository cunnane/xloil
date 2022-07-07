@ECHO OFF

pushd %~dp0

REM Command file for Sphinx documentation

if "%SPHINXBUILD%" == "" (
	set SPHINXBUILD=sphinx-build
)

if "%1" == "" goto help

set XLOIL_SOLN_DIR=%~dp0..

set SOURCEDIR=source
set BUILDDIR=%XLOIL_SOLN_DIR%\build\docs
set PATH=%PATH%;C:\Program Files\doxygen\bin

%SPHINXBUILD% >NUL 2>NUL
if errorlevel 9009 (
	echo.
	echo.The 'sphinx-build' command was not found. Make sure you have Sphinx
	echo.installed, then set the SPHINXBUILD environment variable to point
	echo.to the full path of the 'sphinx-build' executable. Alternatively you
	echo.may add the Sphinx directory to PATH.
	echo.
	echo.If you don't have Sphinx installed, grab it from
	echo.http://sphinx-doc.org/
	exit /b 1
)

mkdir %BUILDDIR%

if "%1" == "doxygen" goto doxygen

if "%1" == "-bin" (
  set XLOIL_BIN_DIR=%XLOIL_SOLN_DIR%\build\%2
  shift
  shift
) else (
  set XLOIL_BIN_DIR=%XLOIL_SOLN_DIR%\build\x64\Debug
)

REM Generate the doc stubs: we can import the core locally, this is
REM really just for ReadTheDocs
echo.Generating doc stubs for xloil_core
python %XLOIL_SOLN_DIR%\libs\xlOil_Python\Package\generate_stubs.py


REM It's very important to pass the -E argument to sphinx, otherwise it does
REM not notice changes to docstrings in python modules and generates the 
REM wrong documentation

REM Set READTHEDOCS so we get consistency with the RTD online version
set READTHEDOCS=1
%SPHINXBUILD% -M %1 %SOURCEDIR% %BUILDDIR% %SPHINXOPTS% -E -W --keep-going %O%
set READTHEDOCS=
goto end


:doxygen

pushd source
REM Need to reverse slashes for doxygen because life is tough
set "XLO_SOLN_DIR=%XLOIL_SOLN_DIR:\=/%"

doxygen xloil.doxyfile
popd
goto end


:help
%SPHINXBUILD% -M help %SOURCEDIR% %BUILDDIR% %SPHINXOPTS% %O%

:end
popd

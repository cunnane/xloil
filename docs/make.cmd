@ECHO OFF

pushd %~dp0

REM Command file for Sphinx documentation

if "%SPHINXBUILD%" == "" (
	set SPHINXBUILD=sphinx-build
)

if "%1" == "" goto help

set XLOIL_SOLN_DIR=%~dp0..
set DOC_BUILD_DIR=%XLOIL_SOLN_DIR%\build\docs
set PY_PACKAGE_DIR=%XLOIL_SOLN_DIR%\libs\xlOil_Python\Package
set SOURCEDIR=%~dp0\source
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

mkdir "%DOC_BUILD_DIR%"

if "%1" == "doxygen" goto doxygen

if "%1" == "-bin" (
  set XLOIL_BIN_DIR=%XLOIL_SOLN_DIR%\build\%2
  shift
  shift
)

REM Generate the doc stubs: since we can import the core pyd locally, this is
REM really just for ReadTheDocs
if exist "%PY_PACKAGE_DIR%\generate_stubs.py" (
	echo.Generating doc stubs for xloil_core
	python "%PY_PACKAGE_DIR%\generate_stubs.py"
)

REM It's very important to pass the -E argument to sphinx, otherwise it does
REM not notice changes to docstrings in python modules and generates the 
REM wrong documentation

REM Set READTHEDOCS so we get consistency with the RTD online version
set READTHEDOCS=1
%SPHINXBUILD% -M %1 "%SOURCEDIR%" "%DOC_BUILD_DIR%" %SPHINXOPTS% -E -W --keep-going %O%
set READTHEDOCS=
goto end


:doxygen

pushd source

REM For some reason doxygen can't create directories itself
mkdir "%DOC_BUILD_DIR%\doxygen\html\doxygen"
mkdir "%DOC_BUILD_DIR%\doxygen\xml\doxygen"

REM Need to reverse slashes for doxygen because life is tough
set "XLO_SOLN_DIR=%XLOIL_SOLN_DIR:\=/%"
doxygen xloil.doxyfile
popd
goto end


:help
%SPHINXBUILD% -M help %SOURCEDIR% %DOC_BUILD_DIR% %SPHINXOPTS% %O%

:end
popd

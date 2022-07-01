@ECHO OFF

pushd %~dp0

REM Command file for Sphinx documentation

if "%SPHINXBUILD%" == "" (
	set SPHINXBUILD=sphinx-build
)

if "%1" == "" goto help

set XLOIL_SOLN_DIR=%~dp0..


set SOURCEDIR=source
set BUILDDIR=%SOLN_DIR%\build\docs
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

if "%1" == "doxygen" goto doxygen

if "%1" == "-bin" (
  set XLOIL_BIN_DIR=%SOLN_DIR%\build\%2
  shift
  shift
) else (
  set XLOIL_BIN_DIR=%SOLN_DIR%\build\x64\Debug
)

REM It's very important to pass the -E argument to sphinx, otherwise it does
REM not notice changes to docstrings in python modules and generates the 
REM wrong documentation

%SPHINXBUILD% -M %1 %SOURCEDIR% %BUILDDIR% %SPHINXOPTS% -E -W --keep-going %O%
goto end

:doxygen
pushd source
set "XLO_SOLN_DIR=%SOLN_DIR:\=/%"

doxygen xloil.doxyfile
popd
goto end


:help
%SPHINXBUILD% -M help %SOURCEDIR% %BUILDDIR% %SPHINXOPTS% %O%

:end
popd

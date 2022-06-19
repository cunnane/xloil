@ECHO OFF

pushd %~dp0

REM Command file for Sphinx documentation

if "%SPHINXBUILD%" == "" (
	set SPHINXBUILD=sphinx-build
)

if "%1" == "" goto help

REM TODO: Build flavour hard coded here!
set XLOIL_SOLN_DIR=%~dp0..
set XLOIL_BIN_DIR=%SOLN_DIR%\build\x64\Debug

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

%SPHINXBUILD% -M %1 %SOURCEDIR% %BUILDDIR% %SPHINXOPTS% -W --keep-going %O%
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

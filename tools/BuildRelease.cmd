set TOOLS_DIR=%~dp0
set SOLN_DIR=%TOOLS_DIR%\..

set ARCH=%1

call "c:\program files (x86)\microsoft visual studio\2017\community\Common7\Tools\vsdevcmd.bat" -no_logo

pushd %SOLN_DIR%

devenv xlOil.sln /Build "Release|%ARCH%"

popd
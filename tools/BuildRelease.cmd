set TOOLS_DIR=%~dp0
set SOLN_DIR=%TOOLS_DIR%\..

call "c:\program files (x86)\microsoft visual studio\2017\community\Common7\Tools\vsdevcmd.bat" -no_logo

pushd %SOLN_DIR%

devenv xlOil.sln /Build "Release|x64"

popd
set TOOLS_DIR=%~dp0
set SOLN_DIR=%TOOLS_DIR%\..

set ARCH=%1

call "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\Tools\vsdevcmd.bat" -no_logo

pushd %SOLN_DIR%

msbuild xlOil.sln /p:Configuration=Release /p:Platform=%ARCH%

popd
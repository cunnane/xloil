set SOLN_DIR=%~dp0\..

pushd %SOLN_DIR%\external

git clone --single-branch --depth=1 --branch=2.10.0-xloil git@github.com:cunnane/pybind11.git pybind11
git clone --single-branch --depth=1 --branch=v1.3.3 https://github.com/marzer/tomlplusplus.git tomlplusplus
git clone --single-branch --depth=1 --branch=v1.4.2 https://github.com/gabime/spdlog.git spdlog
git clone --single-branch --depth=1 --branch=v.0.0.2 https://github.com/vit-vit/CTPL.git ctpl
git clone --single-branch --depth=1 --branch=documentation-stubs https://github.com/cunnane/pybind11-stubgen.git pybind11-stubgen
git clone --single-branch --depth=1 --branch=master https://github.com/tresorit/rdcfswatcherexample.git rdcfswatcher

REM Cloning sqlite is about 100mb and it's much easier to work with the source amalgamation
curl.exe --output sqlite.zip --url https://www.sqlite.org/2022/sqlite-amalgamation-3390300.zip
tar -xf sqlite.zip
rename sqlite-amalgamation-3390300 sqlite

REM asmjit doesn't appear to have versioning, so we keep with our copy of the source
git clone https://github.com/asmjit/asmjit.git asmjit
pushd asmjit
git reset --hard ac77dfcd7
popd

popd

echo.You need to unpack %SOLN_DIR%\external\python.7z
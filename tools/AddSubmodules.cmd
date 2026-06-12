
set SOLN_DIR=%~dp0\..

pushd %SOLN_DIR%\external

rem winreg
git submodule add --name winreg https://github.com/GiovanniDicanio/WinReg winreg
pushd winreg
git checkout v6.1.0
popd

rem pybind11
git submodule add --name pybind11 https://github.com/cunnane/pybind11.git pybind11
pushd pybind11
git checkout v2.13-xloil
popd

rem tomlplusplus
git submodule add --name tomlplusplus https://github.com/marzer/tomlplusplus.git tomlplusplus
pushd tomlplusplus
git checkout v3.4.0
popd

rem spdlog
git submodule add --name spdlog https://github.com/gabime/spdlog.git spdlog
pushd spdlog
git checkout v1.8.5
popd

rem ctpl
git submodule add --name ctpl https://github.com/vit-vit/CTPL.git ctpl
pushd ctpl
git checkout v.0.0.2
popd

rem pybind11-stubgen
git submodule add --name pybind11-stubgen https://github.com/cunnane/pybind11-stubgen.git pybind11-stubgen
pushd pybind11-stubgen
git checkout documentation-stubs
popd

REM asmjit doesn't appear to have versioning, so we keep with our copy of the source
git submodule add --name asmjit https://github.com/asmjit/asmjit.git asmjit
pushd asmjit
git checkout ac77dfcd7
popd

popd

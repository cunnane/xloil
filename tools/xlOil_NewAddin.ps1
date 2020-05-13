
$TargetPath=$args[0]

if ($TargetPath -eq "") {
    Write-Host "Syntax: ", $PSCommandPath, " <AddinName>.xll"
}

$TargetIni = [io.path]::ChangeExtension($TargetPath, "ini")

Copy-Item -Path (Join-Path $PSScriptRoot "xlOil.xll") -Destination $TargetPath 
Copy-Item -Path (Join-Path $PSScriptRoot "NewAddin.ini") -Destination $TargetIni

Write-Host "New addin created at: ", $TargetPath
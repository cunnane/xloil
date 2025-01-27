<#
.SYNOPSIS
Removes all xlOil registry entries

.DESCRIPTION
xlOil should remove its registry entries on exit, however if Excel does not exit cleanly or other issues
disrupt the removal, some keys may remain. This script will not remove entries for any custom COM addins.
#>

$XloilClassIDs = (Get-ChildItem "Registry::HKEY_CLASSES_ROOT\CLSID\*\InprocServer32") `
    | Where-Object {$_.GetValue("") -like "*xlOil.dll"} `
    | ForEach-Object {$_."PSParentPath"}
$XloilClasses = (Get-ChildItem "Registry::HKEY_CLASSES_ROOT") `
    | Where-Object {$_.Name -like "*xlOil*"} `
    | ForEach-Object {$_."PSPath"}
$XloilAddins = (Get-ChildItem HKCU:\Software\Microsoft\Office\Excel\AddIns) `
    | Where-Object {$_.Name -like "*xloil*"} `
    | ForEach-Object {$_."PSPath"}
$XloilAddinData = (Get-ChildItem HKCU:\Software\Microsoft\Office\Excel\AddinsData) `
    | Where-Object {$_.Name -like "*xlOil*"} `
    | ForEach-Object {$_."PSPath"}

$xloilKeys = $XloilClassIDs + $XloilClasses + $XloilAddins + $XloilAddinData


Write-Host "=== Found the following xloil keys:"
$xloilKeys -join "`r`n" | Out-String

$confirmation = Read-Host "=== Press Y to delete"
if ($confirmation -eq 'y') {
    $xloilKeys | ForEach-Object {Remove-Item $_ -Recurse}
}

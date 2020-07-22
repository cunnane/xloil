
function Remove-Addin {
    param ([string]$AddinPath, $Version)
    $RegKey=(gi "HKCU:\Software\Microsoft\Office\${Version}\Excel\Add-in Manager" -ErrorAction SilentlyContinue)

    if ($RegKey -eq $NULL) {exit}

    # Cycles through all the properties and deletes it if it contains the file name.
    foreach ($prop in $RegKey.Property) {
        if ($prop.Contains($AddinPath)) {
            $RegKey|Remove-ItemProperty -name $prop
        }
    }
}

#####################################################################################################
#
# Script Start
#

$ADDIN_NAME   = "xlOil.xll"
$INIFILE_NAME = "xlOil.ini"

$XloilAppData = Join-Path $env:APPDATA "xlOil"

#
# Start Excel to get some environment settings (could probably get
# them from the registry more quickly...)
#
$Excel = New-Object -Com Excel.Application
$ExcelVersion = $Excel.Version
$XlStartPath = $Excel.StartupPath
$Excel.Quit()
$Excel = $null

$AddinPath = Join-Path $XlStartPath $ADDIN_NAME

# Ensure no xlOil addins are in the registry
Remove-Addin $AddinPath $ExcelVersion

Remove-Item –path $AddinPath 
Write-Host (Join-Path $AddinPath $ADDIN_NAME), "removed"
Write-Host "Left settings files in ", $XloilAppData 

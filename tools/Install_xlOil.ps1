
$ADDIN_NAME = "xlOil.xll"
$OurAppData = Join-Path $env:APPDATA "xlOil"

function Remove-From-Resiliancy {
    param ([string]$FileName, [string]$OfficeVersion)

    # Source https://stackoverflow.com/questions/751048/how-to-programatically-re-enable-documents-in-the-ms-office-list-of-disabled-fil

    #Converts the File Name string to UTF16 Hex
    $FileName_UniHex=""
    [System.Text.Encoding]::ASCII.GetBytes($FileName.ToLower()) | %{$FileName_UniHex+="{0:X2}00" -f $_}

    #Tests to see if the Disabled items registry key exists
    $RegKey=(gi "HKCU:\Software\Microsoft\Office\${OfficeVersion}\Excel\Resiliency\DisabledItems\")
    if ($RegKey -eq $NULL) {exit}

    #Cycles through all the properties and deletes it if it contains the file name.
    foreach ($prop in $RegKey.Property) {
        $Val=""
        ($RegKey|gp).$prop | %{$Val+="{0:X2}" -f $_}
        if ($Val.Contains($FileName_UniHex)) {
            $RegKey|Remove-ItemProperty -name $prop
        }
    }
}

$Excel = New-Object -Com Excel.Application
$ExcelVersion = $Excel.Version

# Just in case we got put in Excel's naughty corner for misbehaving addins
Remove-From-Resiliancy $ADDIN_NAME $ExcelVersion

# You can't add an add-in unless there's an open and visible workbook.
# It's a long-standing Excel bug which, like so many others, Microsoft
# is unlikely to fix, not whilst the important task of tweaking the UI
# appearance with every Office version takes priority.

$Excel.Visible = $true
$Worbook = $Excel.Workbooks.Add()
$Worbook.Sheets(1).Cells(1,1).Value = "Instaling xlOil addin"
$AddinPath = Join-Path $PSScriptRoot $ADDIN_NAME
$Addin = $Excel.AddIns.Add($AddinPath)
$Addin.Installed = $true
$Worbook.Close($false)
$Excel.quit()

Copy-Item -path $PSScriptRoot -include "*.ini"  -Destination $OurAppData

Write-Host $AddinPath, "installed"
Write-Host "Settings files placed in ",$OurAppData


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

#####################################################################################################
#
# Script Start
#

$ADDIN_NAME = "xlOil.xll"
$INIFILE_NAME = "xlOil.ini"
$OurAppData = Join-Path $env:APPDATA "xlOil"

#
# Start Excel
#
$Excel = New-Object -Com Excel.Application
$ExcelVersion = $Excel.Version

# Just in case we got put in Excel's naughty corner for misbehaving addins
Remove-From-Resiliancy $ADDIN_NAME $ExcelVersion

#
# You can't add an add-in unless there's an open and visible workbook.
# It's a long-standing Excel bug which, like so many others, Microsoft
# is unlikely to fix, not whilst the important task of tweaking the UI
# appearance with every Office version takes priority.
#
$Excel.Visible = $true
$Workbook = $Excel.Workbooks.Add()
$Worksheet = $Workbook.Sheets(1)
$Worksheet.Cells(1,1).Value = "Instaling xlOil addin"
$AddinPath = Join-Path $PSScriptRoot $ADDIN_NAME
$Addin = $Excel.AddIns.Add($AddinPath)
$Addin.Installed = $true
$Workbook.Close($false)

#
# We need to null all the COM refs we used or Excel won't actually quit
# even after this script has ended. It's a well-known problem see for example
# https://stackoverflow.com/questions/42113082/excel-application-object-quit-leaves-excel-exe-running
#
$Workbook = $null
$Worksheet = $null
$Addin = $null
$Excel.Quit()
$Excel = $null

#
# Copy the ini file to APPDATA
#
$IniFile = (Join-Path $OurAppData $INIFILE_NAME)

if (!(Test-Path -Path $IniFile -PathType leaf)) {

	mkdir -Force $OurAppData | Out-Null
	Copy-Item -path (Join-Path $PSScriptRoot $INIFILE_NAME) -Destination $IniFile

	#
	# Set the PATH environment in the ini so we can found xloil.dll if required
	#
	(Get-Content -Encoding UTF8 -Path $IniFile) `
		-replace "'''%PATH%'''", "'''%PATH%;$PSScriptRoot'''" |
	  Out-File -Encoding UTF8 $IniFile 

	Write-Host "Settings files placed in ",$OurAppData

} else {

	Write-Host "Found existing settings file at `n",$IniFile , "`nCheck [Environment] block points to `n", $AddinPath

}


Write-Host $AddinPath, "installed"

#
# Helps ensure Excel really closes when the script exits
# 
[system.gc]::Collect()

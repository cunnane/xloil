
function Remove-From-Resiliancy {
    param ([string]$FileName, [string]$OfficeVersion)

    # Source https://stackoverflow.com/questions/751048/how-to-programatically-re-enable-documents-in-the-ms-office-list-of-disabled-fil

    #Converts the File Name string to UTF16 Hex
    $FileName_UniHex=""
    [System.Text.Encoding]::ASCII.GetBytes($FileName.ToLower()) | %{$FileName_UniHex+="{0:X2}00" -f $_}

    #Tests to see if the Disabled items registry key exists
    $RegKey=(gi "HKCU:\Software\Microsoft\Office\${OfficeVersion}\Excel\Resiliency\DisabledItems\" -ErrorAction SilentlyContinue)
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

function Check-VBA-Access {
	param ([string]$OfficeVersion)

	$UserKey=(Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\${OfficeVersion}\Excel\Security" -ErrorAction SilentlyContinue)
	$MachineKey=(Get-ItemProperty -Path "HKLM:\Software\Microsoft\Office\${OfficeVersion}\Excel\Security" -ErrorAction SilentlyContinue)

	if (($MachineKey.AccessVBOM -eq 0) -or ($UserKey.AccessVBOM -eq 0)) 
	{ 
		Write-Host (
			"To ensure xlOil local functions work, allow access to the VBA Object Model in`n" +
			"Excel > File > Options > Trust Center > Trust Center Settings > Macro Settings`n")
	}
}

#####################################################################################################
#
# Script Start
#

$ADDIN_NAME   = "xlOil.xll"
$INIFILE_NAME = "xlOil.ini"

$XloilAppData = Join-Path $env:APPDATA "xlOil"
$AddinPath = Join-Path $PSScriptRoot $ADDIN_NAME

#
# Start Excel to get some environment settings (could probably get
# them from the registry more quickly...)
#
$Excel = New-Object -Com Excel.Application
$ExcelVersion = $Excel.Version
$XlStartPath = $Excel.StartupPath
$Excel.Quit()
$Excel = $null

# Just in case we got put in Excel's naughty corner for misbehaving addins
Remove-From-Resiliancy $ADDIN_NAME $ExcelVersion

# Check access to the VBA Object model (for local functions)
Check-VBA-Access $ExcelVersion

# Ensure XLSTART dir really exists
mkdir -Force $XlStartPath | Out-Null

# Copy the XLL

Copy-Item -path (Join-Path $PSScriptRoot $ADDIN_NAME) -Destination $XlStartPath

# Copy the ini file to APPDATA, avoiding overwritting any existing ini
md $XloilAppData -ErrorAction Ignore
$IniFile = (Join-Path $XloilAppData $INIFILE_NAME)
if (!(Test-Path -Path $IniFile -PathType leaf)) {
	
	Copy-Item -path (Join-Path $PSScriptRoot $INIFILE_NAME) -Destination $IniFile

} else {

	Write-Host ("Found existing settings file at `n", $IniFile)

}
#
# Set the PATH environment in the ini so we can find xlOil.dll
#
(Get-Content -Encoding UTF8 -Path $IniFile) `
	-replace "^(\s*XLOIL_PATH\s*=).*", "`$1'''$PSScriptRoot'''" |
	Out-File -Encoding UTF8 $IniFile 

Write-Host "Edited settings file ", $IniFile 

Write-Host "$ADDIN_NAME installed from $PSScriptRoot to $XlStartPath `n"

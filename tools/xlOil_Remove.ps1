
$ADDIN_NAME = "xlOil.xll"
$OurAppData = Join-Path $env:APPDATA "xlOil"

function Remove-Addin {
    param ([string]$AddinPath, $Version)
    $RegKey=(gi "HKCU:\Software\Microsoft\Office\${Version}\Excel\Add-in Manager")

    if ($RegKey -eq $NULL) {exit}

    #Cycles through all the properties and deletes it if it contains the file name.
    foreach ($prop in $RegKey.Property) {
        if ($prop.Contains($AddinPath)) {
            $RegKey|Remove-ItemProperty -name $prop
        }
    }

}

$Excel = New-Object -Com Excel.Application
$ExcelVersion = $Excel.Version
$Excel.Visible = $true

# You can't add an add-in unless there's an open and visible workbook.
$Worbook = $Excel.Workbooks.Add()
$Worbook.Sheets(1).Cells(1,1).Value2 = "Removing xlOil addin"

# For some reason I can't get the addins collection to index
# from a string like in normal COM, so we have to loop
$AddinPath = ""
For ($i = 1; $i -le $Excel.AddIns.Count; $i++) {
    $Addin = $Excel.AddIns[$i]
    if ($Addin.Name -like $ADDIN_NAME) {
        $Addin.Installed = $false
        $AddinPath = $Addin.Path
    }
}

$Worbook.Close($false)
$Excel.quit()

# Microsoft neglected to provide a Remove function on the add-ins 
# collection, because where would all the fun be in programming
# if they did all the work?

Remove-Addin $AddinPath $ExcelVersion

Write-Host (Join-Path $AddinPath $ADDIN_NAME), "removed"
Write-Host "Left settings files in ",$OurAppData 

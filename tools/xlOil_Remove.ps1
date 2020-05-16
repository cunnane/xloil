
$ADDIN_NAME = "xlOil.xll"
$OurAppData = Join-Path $env:APPDATA "xlOil"

function Remove-Addin {
    param ([string]$AddinPath, $Version)
    $RegKey=(gi "HKCU:\Software\Microsoft\Office\${Version}\Excel\Add-in Manager")

    if ($RegKey -eq $NULL) {exit}

    # Cycles through all the properties and deletes it if it contains the file name.
    foreach ($prop in $RegKey.Property) {
        if ($prop.Contains($AddinPath)) {
            $RegKey|Remove-ItemProperty -name $prop
        }
    }
}

$Excel = New-Object -Com Excel.Application
$ExcelVersion = $Excel.Version
$Excel.Visible = $true

#
# You can't add an add-in unless there's an open and visible workbook.
# There's no logical reason for this, it's a bug.
#
$Worbook = $Excel.Workbooks.Add()
$Worbook.Sheets(1).Cells(1,1).Value2 = "Removing xlOil addin"


#
# For some reason I can't get the addins collection to index
# from a string like in normal COM, so we have to loop
#
$AddinPath = ""
For ($i = 1; $i -le $Excel.AddIns.Count; $i++) {
    $Addin = $Excel.AddIns[$i]
    if ($Addin.Name -like $ADDIN_NAME) {
        $Addin.Installed = $false
        $AddinPath = $Addin.Path
    }
}
$Worbook.Close($false)

#
# We need to null all the COM refs we used or Excel won't actually quit
# even after this script has ended. It's a well-known problem see for example
# https://stackoverflow.com/questions/42113082/excel-application-object-quit-leaves-excel-exe-running
#
$Workbook = $null
$Addin = $null
$Excel.Quit()
$Excel = $null

#
# Microsoft neglected to provide a Remove function on the add-ins 
# collection, because where would all the fun be in programming if
# they wrote a fully featured API which works as expected and documented?
#
Remove-Addin $AddinPath $ExcelVersion

Write-Host (Join-Path $AddinPath $ADDIN_NAME), "removed"
Write-Host "Left settings files in ",$OurAppData 

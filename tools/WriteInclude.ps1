
$HeaderDir = $args[0]
$TargetDir = $args[1]


# Copy-Item -Path $xlOilDir -Filter *.h -Destination $TargetDir


$xloilHeader = Join-Path $TargetDir "xloil.h"
Set-Content -Path $xloilHeader -Value "#pragma once"

Get-ChildItem -Path $HeaderDir -File -Filter *.h -Name | 
    Foreach-Object {
        return "#include `"" + $_ + "`""
    } |
    Out-File $xloilHeader -Append utf8

$HeaderDir = $args[0]
$TargetDir = $args[1]


$xloilHeader = Join-Path $TargetDir "xlOil.h"
Set-Content -Path $xloilHeader -Value "#pragma once"

Get-ChildItem -Path $HeaderDir -File -Filter *.h -Name | 
	where {$_ -notmatch 'xlOil.h'} |
    Foreach-Object {
        return "#include `"" + $_ + "`""
    } |
    Out-File $xloilHeader -Append utf8
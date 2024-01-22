﻿
$HeaderDir = $args[0]
$TargetDir = $args[1]

$xloilHeader = Join-Path $TargetDir "xlOil.h"
New-Item $xloilHeader -ErrorAction SilentlyContinue
Set-Content -Path $xloilHeader -Value "#pragma once"

Get-ChildItem -Path $HeaderDir -File -Filter *.h -Name | 
	where {$_ -NotMatch 'xlOil.h|XllEntryPoint.h|WindowsSlim.h|ExcelTypeLib.h'} |
    Foreach-Object {
        return "#include `"" + $_ + "`""
    } |
    Out-File $xloilHeader -Append utf8
$Match = Select-String -path "include\xloil\Version.h" -pattern "_VERSION +(\d)" -AllMatches | Foreach-Object {$_.Matches}
Set-Content -Path 'Version.txt' -Value "$($Match.Groups[1]).$($Match.Groups[3]).$($Match.Groups[5])"

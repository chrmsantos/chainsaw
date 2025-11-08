. $PSScriptRoot\Helpers.ps1
$v = Get-VbaFiles
Write-Host "Count: $($v.Count)"
foreach ($f in $v) { Write-Host $f.FullName }
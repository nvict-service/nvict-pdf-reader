param(
    [string]$Version,
    [string]$Year
)

$file = "version_info.txt"
if (-not (Test-Path $file)) { exit 0 }

$parts = $Version.Split('.')
$maj = [int]$parts[0]
$min = if ($parts.Count -gt 1) { [int]$parts[1] } else { 0 }
$fullVer = "$Version.0.0"

$t = Get-Content $file -Raw -Encoding UTF8
$t = $t -replace 'filevers=\(\d+, \d+, \d+, \d+\)', "filevers=($maj, $min, 0, 0)"
$t = $t -replace 'prodvers=\(\d+, \d+, \d+, \d+\)', "prodvers=($maj, $min, 0, 0)"
$t = $t -replace "FileVersion', u'\d+\.\d+\.\d+\.\d+'", "FileVersion', u'$fullVer'"
$t = $t -replace "ProductVersion', u'\d+\.\d+\.\d+\.\d+'", "ProductVersion', u'$fullVer'"
$t = $t -replace 'Copyright \d{4}', "Copyright $Year"
Set-Content $file -Value $t -Encoding UTF8 -NoNewline

Write-Host "[OK] version_info.txt bijgewerkt: v$fullVer, Copyright $Year"

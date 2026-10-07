# ORT lab management system - XP lite client - build a hand-over zip.
#
# ASCII-ONLY ON PURPOSE (see tools/build_xp.ps1 for the reason).
#
# Takes the already built onedir bundle (dist\ORT-XP), zips it as
# dist\ORT-XP-<version>-<date>.zip, writes a SHA256 file next to it and prints
# both, so the hand-over has a verifiable artifact.
#
# Usage:
#   .\tools\make_release.ps1
#   .\tools\make_release.ps1 -Version 0.4.0

param(
    [string]$Version = "",
    [string]$Bundle = ""
)

$ErrorActionPreference = "Continue"
$root = Split-Path -Parent $PSScriptRoot
if (-not $Bundle) { $Bundle = Join-Path $root "dist\ORT-XP" }

if (-not (Test-Path $Bundle)) {
    throw "bundle not found: $Bundle - run tools\build_xp.ps1 first"
}

if (-not $Version) {
    # read VERSION = "x.y.z" from ort_xp/version.py (ASCII file, safe to parse textually)
    $versionFile = Join-Path $root "ort_xp\version.py"
    $match = Select-String -Path $versionFile -Pattern '^VERSION\s*=\s*"([^"]+)"' | Select-Object -First 1
    if (-not $match) { throw "cannot read VERSION from $versionFile" }
    $Version = $match.Matches[0].Groups[1].Value
}

$stamp = Get-Date -Format "yyyyMMdd"
$zipName = "ORT-XP-$Version-$stamp.zip"
$zipPath = Join-Path (Split-Path -Parent $Bundle) $zipName

if (Test-Path $zipPath) { Remove-Item $zipPath -Force -ErrorAction Stop }
Write-Host "zipping $Bundle" -ForegroundColor Cyan
Compress-Archive -Path (Join-Path $Bundle "*") -DestinationPath $zipPath -CompressionLevel Optimal -ErrorAction Stop

$hash = Get-FileHash $zipPath -Algorithm SHA256
$sizeMb = (Get-Item $zipPath).Length / 1MB
$hashFile = "$zipPath.sha256.txt"
"$($hash.Hash)  $zipName" | Set-Content -Path $hashFile -Encoding ASCII

# also record what is inside, for the hand-over record
$fileCount = (Get-ChildItem $Bundle -Recurse -File | Measure-Object).Count
$extras = (Get-ChildItem $Bundle -File | Where-Object { $_.Extension -in @(".cmd", ".vbs") } | Select-Object -ExpandProperty Name) -join ", "

Write-Host ""
Write-Host ("release : {0}" -f $zipPath) -ForegroundColor Green
Write-Host ("size    : {0:N1} MB ({1} files in the bundle)" -f $sizeMb, $fileCount) -ForegroundColor Green
Write-Host ("sha256  : {0}" -f $hash.Hash) -ForegroundColor Green
Write-Host ("hash    : {0}" -f $hashFile) -ForegroundColor Green
Write-Host ("extras  : {0}" -f $(if ($extras) { $extras } else { "(none)" })) -ForegroundColor Green

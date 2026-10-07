# ORT lab management system - XP lite client - alternating read/write verification.
#
# ASCII-ONLY ON PURPOSE: this file is written in plain ASCII so it parses the same under
# Windows PowerShell (ANSI/cp950 default) and PowerShell 7. Chinese notes live in
# docs/02-data-contract.md; the XP-side steps are implemented in tools/xp_step.py.
#
# WHAT IT PROVES
#   The one shared ort_plans.db survives being written by BOTH ends while both keep a
#   connection open:
#     - the XP client (python sqlite3, same code path as its UI), and
#     - the main program's own provider (System.Data.SQLite - the ADO.NET provider FreeSql
#       sits on, loaded from bin\Debug),
#   with no corruption, no permanent lock, and each end seeing the other end's committed rows.
#   It also checks the lock path on purpose: while the .NET side holds a write transaction,
#   the XP write must fail CLEANLY (database is locked) and must succeed again right after the
#   rollback - that is the "no permanent lock" part.
#
# The real database is NEVER touched: everything runs on a temporary copy (db + wal + shm).
#
# Usage:
#   .\tools\verify_alternating.ps1
#   .\tools\verify_alternating.ps1 -Python "D:\Python34-32\python.exe"
#   .\tools\verify_alternating.ps1 -Keep          # keep the temp copy for inspection
#
# Exit code 0 = all checks passed, 1 = at least one check failed.

param(
    [string]$Python = "python",
    [string]$SourceDb = "",
    [switch]$Keep
)

$ErrorActionPreference = "Continue"
$root = Split-Path -Parent $PSScriptRoot
if (-not $SourceDb) { $SourceDb = Join-Path $root "..\bin\Debug\Data\ort_plans.db" }
$driverDir = Join-Path $root "..\bin\Debug"
$stepScript = Join-Path $root "tools\xp_step.py"
if (-not (Test-Path $stepScript)) { throw "step helper not found: $stepScript" }

if (-not (Test-Path $SourceDb)) { throw "source database not found: $SourceDb" }
if (-not (Test-Path (Join-Path $driverDir "System.Data.SQLite.dll"))) {
    throw "System.Data.SQLite.dll not found under $driverDir (build the main program first)"
}

$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$work = Join-Path $env:TEMP "ort-alt-verify-$stamp"
New-Item -ItemType Directory -Path $work -Force | Out-Null
Copy-Item $SourceDb (Join-Path $work "ort_plans.db") -Force
foreach ($suffix in @("-wal", "-shm")) {
    $extra = "$SourceDb$suffix"
    if (Test-Path $extra) { Copy-Item $extra (Join-Path $work "ort_plans.db$suffix") -Force }
}
Write-Host "work copy: $work" -ForegroundColor Cyan
$dbPath = Join-Path $work "ort_plans.db"

$results = New-Object System.Collections.ArrayList
function Check([string]$name, [bool]$ok, [string]$detail) {
    [void]$results.Add([pscustomobject]@{ Name = $name; Ok = $ok; Detail = $detail })
    $colour = if ($ok) { "Green" } else { "Red" }
    Write-Host ("  [{0}] {1}{2}" -f $(if ($ok) { "OK" } else { "!!" }), $name, $(if ($detail) { " -> $detail" } else { "" })) -ForegroundColor $colour
}

# --------------------------------------------------------------------------- drivers

Add-Type -Path (Join-Path $driverDir "System.Data.SQLite.dll")
$net = New-Object System.Data.SQLite.SQLiteConnection("Data Source=$dbPath;Version=3;")
$net.Open()
Write-Host "main-program driver open (System.Data.SQLite $($net.ServerVersion))" -ForegroundColor Cyan

function Sql-Scalar([string]$sql) {
    $cmd = $net.CreateCommand()
    $cmd.CommandText = $sql
    try { return $cmd.ExecuteScalar() } finally { $cmd.Dispose() }
}
function Sql-Exec([string]$sql) {
    $cmd = $net.CreateCommand()
    $cmd.CommandText = $sql
    try { return $cmd.ExecuteNonQuery() } finally { $cmd.Dispose() }
}
function Invoke-Xp([string[]]$Arguments) {
    # Run from the xp-client root with a RELATIVE script path on purpose: Python 3.4 uses
    # the ANSI code page (mbcs) for some import/linecache paths, and this repository lives
    # under a path with Simplified Chinese ("ORT...") that cp950 cannot represent - an
    # absolute script path there fails with "ImportError: No module named 'ort_xp'".
    Push-Location $root
    try {
        $output = & $Python "tools\xp_step.py" --data-folder $work @Arguments 2>&1
        $text = ($output | Out-String).Trim()
    }
    finally {
        Pop-Location
    }
    return $text
}
function Now-Text() { return (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") }

# --------------------------------------------------------------------------- steps

Write-Host "== 1. main-program driver writes a plan, XP client reads it ==" -ForegroundColor Cyan
$plansBefore = [int](Sql-Scalar "SELECT COUNT(*) FROM plans")
$dbStamp = Now-Text
$null = Sql-Exec ("INSERT INTO plans (JobNo, ModelName, TestItem, Stage, Owner, Status, StartDate, EndDate, Remark, CreatedBy, CreatedAt, UpdatedBy, UpdatedAt) " +
    "VALUES ('ALT-VERIFY-01', 'ALTMODEL', 'ALTITEM', 'DVT', 'ALT', 'Ongoing', '2026-09-01 00:00:00', '2026-12-31 00:00:00', 'ALTVERIFY-REMARK', 'alt-verify', '$dbStamp', 'alt-verify', '$dbStamp')")
Check ".NET INSERT plan ALT-VERIFY-01" ((Sql-Scalar "SELECT COUNT(*) FROM plans") -eq ($plansBefore + 1)) "plans $plansBefore -> $($plansBefore + 1)"
Check "XP sees the plan (while .NET connection stays open)" ((Invoke-Xp @("plan-exists", "ALT-VERIFY-01")) -eq "yes")
Check "XP reads the .NET-written remark" ((Invoke-Xp @("plan-remark", "ALT-VERIFY-01")) -eq "ALTVERIFY-REMARK")

Write-Host "== 2. XP client writes, main-program driver reads ==" -ForegroundColor Cyan
$newResult = Invoke-Xp @("new-requisition", "ALT2609001")
Check "XP INSERT requisition ALT2609001" ($newResult -like "ok *") $newResult
Check ".NET sees the XP-written requisition" ((Sql-Scalar "SELECT COUNT(*) FROM requisitions WHERE RequisitionNo = 'ALT2609001'") -eq 1)
$logCount = [int](Sql-Scalar "SELECT COUNT(*) FROM plan_change_logs WHERE Summary LIKE '%ALT2609001%'")
Check ".NET sees the XP-written change log" ($logCount -ge 1) "rows=$logCount"
$xpUpdate = Invoke-Xp @("update-requisition", "ALT2609001", "XP-UPDATED-REMARK")
Check "XP UPDATE requisition remark" ($xpUpdate -eq "updated") $xpUpdate
Check ".NET reads the XP update" ((Sql-Scalar "SELECT Remark FROM requisitions WHERE RequisitionNo = 'ALT2609001'") -eq "XP-UPDATED-REMARK")

Write-Host "== 3. .NET updates, XP client reads the update ==" -ForegroundColor Cyan
$null = Sql-Exec "UPDATE plans SET Status = 'Close', UpdatedAt = '$dbStamp' WHERE JobNo = 'ALT-VERIFY-01'"
Check "XP sees the .NET status update (Close)" ((Invoke-Xp @("plan-status-is-close", "ALT-VERIFY-01")) -eq "yes")

Write-Host "== 4. lock behaviour: .NET holds a write transaction ==" -ForegroundColor Cyan
$tx = $net.BeginTransaction()
$cmd = $net.CreateCommand()
$cmd.Transaction = $tx
$cmd.CommandText = "INSERT INTO plans (JobNo, ModelName, Status, CreatedBy, CreatedAt) VALUES ('ALT-LOCK-01', 'ALTMODEL', 'Ongoing', 'alt-verify', '$dbStamp')"
$null = $cmd.ExecuteNonQuery()
$cmd.Dispose()
Check "XP can still READ while .NET holds the write lock" ((Invoke-Xp @("plan-exists", "ALT-VERIFY-01")) -eq "yes")
$locked = Invoke-Xp @("--timeout", "3", "--retries", "0", "new-requisition", "ALT2609002")
Check "XP write fails CLEANLY while locked (no hang, no corruption)" ($locked -like "locked*") (($locked -split "`n")[0])
$tx.Rollback()
$tx.Dispose()
Check "uncommitted .NET row is gone after rollback" ((Sql-Scalar "SELECT COUNT(*) FROM plans WHERE JobNo = 'ALT-LOCK-01'") -eq 0)
$retry = Invoke-Xp @("new-requisition", "ALT2609002")
Check "XP write succeeds again after rollback (no permanent lock)" ($retry -like "ok *") $retry
Check ".NET sees the row written after the rollback" ((Sql-Scalar "SELECT COUNT(*) FROM requisitions WHERE RequisitionNo = 'ALT2609002'") -eq 1)

Write-Host "== 5. integrity and durability ==" -ForegroundColor Cyan
Check ".NET integrity_check" ((Sql-Scalar "PRAGMA integrity_check") -eq "ok")
Check "XP integrity_check" ((Invoke-Xp @("integrity")) -eq "ok")
$net.Close()
$xpCount = Invoke-Xp @("plan-count")
Check "XP still reads after .NET closes its connection" ($xpCount -match '^\d+$') "plans=$xpCount"

# --------------------------------------------------------------------------- summary

$failed = @($results | Where-Object { -not $_.Ok })
Write-Host ""
Write-Host ("checked {0} items, failed {1}" -f $results.Count, $failed.Count) -ForegroundColor $(if ($failed.Count -eq 0) { "Green" } else { "Red" })
if ($failed.Count -gt 0) {
    $failed | ForEach-Object { Write-Host ("  FAILED: {0} -> {1}" -f $_.Name, $_.Detail) -ForegroundColor Red }
}

if ($Keep) {
    Write-Host "kept: $work" -ForegroundColor Yellow
} else {
    Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
}

if ($failed.Count -gt 0) { exit 1 }
exit 0

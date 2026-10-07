# ORT lab management system - XP lite client - build the XP bundle (PyInstaller 3.3.1)
#
# ASCII-ONLY ON PURPOSE: this file is written in plain ASCII so it parses the same
# under Windows PowerShell (ANSI/cp950 default) and PowerShell 7. Chinese notes for
# this script live in docs/01-environment.md.
#
# Requires Python 3.4 (32-bit) + PyInstaller 3.3.1 (only 3.3.1-3.4 produce Windows
# bootloaders that run on Windows XP). Install them with tools/install_py34.ps1.
#
# Usage:
#   .\tools\build_xp.ps1 -Python "D:\Python34-32\python.exe"
#   .\tools\build_xp.ps1 -Python "D:\Python34-32\python.exe" -SkipCheck
#
# Output: <xp-client>\dist\ORT-XP\ (onedir; copy the whole folder to the XP machine).

param(
    [string]$Python = "python",
    # Staging/build directory: MUST be an ASCII path, see the note below.
    [string]$StageDir = "$env:TEMP\ort-xp-build",
    [switch]$SkipCheck
)

# Continue (not Stop) on purpose: native commands (python.exe) write INFO/WARNING to
# stderr, and with $ErrorActionPreference='Stop' PowerShell turns that into a
# terminating error and aborts the build. File operations that must fail loudly use
# -ErrorAction Stop, and native results are checked via $LASTEXITCODE.
$ErrorActionPreference = "Continue"
$root = Split-Path -Parent $PSScriptRoot
Push-Location $root
try {
    Write-Host "== 1/7 check interpreter ==" -ForegroundColor Cyan
    & $Python -c "import sys; print('Python %d.%d.%d (%d bit)' % (sys.version_info[0], sys.version_info[1], sys.version_info[2], 64 if sys.maxsize > 2**32 else 32))"
    if ($LASTEXITCODE -ne 0) { throw "cannot run interpreter: $Python" }

    $versionOk = & $Python -c "import sys; print(1 if sys.version_info[:2] == (3, 4) else 0)"
    if ($versionOk.Trim() -ne "1") {
        Write-Warning "Interpreter is not Python 3.4; XP bundles must be built with 3.4 (32-bit)."
        if (-not $SkipCheck) { throw "wrong interpreter version (use -SkipCheck only for troubleshooting)" }
    }

    Write-Host "== 2/7 check PyInstaller ==" -ForegroundColor Cyan
    & $Python -m PyInstaller --version
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller not available; run tools\install_py34.ps1 first" }

    Write-Host "== 3/7 Python 3.4 syntax floor check ==" -ForegroundColor Cyan
    & $Python "tools\compat_check.py" "ort_xp" 2>&1 | ForEach-Object { Write-Host $_ }
    if ($LASTEXITCODE -ne 0) { throw "source contains Python 3.4-incompatible syntax" }

    Write-Host "== 4/7 stage sources to an ASCII path ==" -ForegroundColor Cyan
    # WHY: PyInstaller 3.3.1 reads the spec and writes its warning log using the system
    # ANSI encoding (cp950 here). This repository lives under a path containing
    # non-ASCII characters (ORT...), which makes the build die with
    # UnicodeEncodeError / UnicodeDecodeError. Building from %TEMP% avoids that.
    if (Test-Path $StageDir) { Remove-Item $StageDir -Recurse -Force -ErrorAction Stop }
    New-Item -ItemType Directory -Path $StageDir -Force -ErrorAction Stop | Out-Null
    Copy-Item (Join-Path $root "ort_xp") (Join-Path $StageDir "ort_xp") -Recurse -Force -ErrorAction Stop
    Copy-Item (Join-Path $root "packaging") (Join-Path $StageDir "packaging") -Recurse -Force -ErrorAction Stop
    Write-Host "stage dir: $StageDir"

    Write-Host "== 5/7 build (onedir) ==" -ForegroundColor Cyan
    $buildLog = Join-Path $StageDir "pyinstaller.log"
    Push-Location $StageDir
    try {
        & $Python -m PyInstaller "packaging\ort_xp.spec" --noconfirm --clean --distpath dist --workpath build *> $buildLog
        $pyExit = $LASTEXITCODE
    }
    finally {
        Pop-Location
    }
    if (Test-Path $buildLog) { Get-Content $buildLog -Tail 8 | ForEach-Object { Write-Host "  $_" } }
    if ($pyExit -ne 0) { throw "build failed (full log: $buildLog)" }

    $built = Join-Path $StageDir "dist\ORT-XP"
    if (-not (Test-Path $built)) { throw "no output produced: $built" }

    Write-Host "== 6/7 smoke: run the built exe from the ASCII stage path ==" -ForegroundColor Cyan
    # WHY the stage path: ORT-XP.exe is a windowed app and refuses to run from a path the
    # ANSI code page cannot encode (the repo path here contains Chinese). Running it from
    # the ASCII stage dir proves the bundle loads its own python34/msvcr100/tcl DLLs.
    $proc = Start-Process -FilePath (Join-Path $built "ORT-XP.exe") -ArgumentList "--version" -PassThru -Wait
    if ($proc.ExitCode -ne 0) { throw "built exe failed to start (exit code $($proc.ExitCode))" }
    Write-Host "  ORT-XP.exe --version -> exit 0" -ForegroundColor Green

    Write-Host "== 7/7 copy deployment extras into the bundle ==" -ForegroundColor Cyan
    # extras: selftest.cmd (double-click self test) + make_shortcut.vbs (WSH desktop shortcut;
    # Windows XP has WSH but no PowerShell). ASCII file names and ASCII content on purpose.
    $extras = Join-Path $root "packaging\xp-extras"
    if (Test-Path $extras) {
        Copy-Item (Join-Path $extras "*") $built -Force -ErrorAction Stop
        Get-ChildItem $extras -File | ForEach-Object { Write-Host ("  + {0}" -f $_.Name) }
    }
    else {
        Write-Warning "no deployment extras found at $extras"
    }

    $target = Join-Path $root "dist\ORT-XP"
    if (Test-Path (Join-Path $root "dist")) { Remove-Item (Join-Path $root "dist") -Recurse -Force -ErrorAction Stop }
    New-Item -ItemType Directory -Path (Join-Path $root "dist") -Force -ErrorAction Stop | Out-Null
    Copy-Item $built $target -Recurse -Force -ErrorAction Stop

    $sizeMb = (Get-ChildItem $target -Recurse -File | Measure-Object Length -Sum).Sum / 1MB
    Write-Host ""
    Write-Host ("done: {0} ({1:N1} MB)" -f $target, $sizeMb) -ForegroundColor Green
    Write-Host "deploy: copy the whole ORT-XP folder to the XP machine and run ORT-XP.exe" -ForegroundColor Green
    Write-Host "IMPORTANT: deploy to an ASCII-only path (e.g. C:\ORT-XP)." -ForegroundColor Yellow
    Write-Host "  On Windows the module loader encodes paths with the ANSI code page; a path" -ForegroundColor Yellow
    Write-Host "  containing characters that code page cannot represent aborts startup." -ForegroundColor Yellow
    $encodeOk = & $Python -c "import sys; sys.path.insert(0, '.'); from ort_xp import compat; print(1 if compat.can_encode_path(r'$target') else 0)"
    if ("$encodeOk".Trim() -ne "1") {
        Write-Warning "This path cannot be encoded by the current ANSI code page; always run the app from an ASCII path."
    }}
finally {
    Pop-Location
}

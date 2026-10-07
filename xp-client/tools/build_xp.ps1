# ORT 实验室管理系统 · XP 精简客户端 —— XP 发布包构建脚本
#
# 只能在 Python 3.4（32 位）+ PyInstaller 3.3.1 环境下执行：
# XP 上可用的 PyInstaller 启动器只有 3.3.1–3.4 才生成得出来。
#
# 用法：
#   .\tools\build_xp.ps1 -Python "D:\Python34-32\python.exe"
#   .\tools\build_xp.ps1 -Python "D:\Python34-32\python.exe" -SkipCheck
#
# 产物：<xp-client>\dist\ORT-XP\  （onedir，整个目录拷到 XP 机器即可用）

param(
    [string]$Python = "python",
    [switch]$SkipCheck
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Push-Location $root
try {
    Write-Host "== 1/4 检查解释器版本 ==" -ForegroundColor Cyan
    & $Python -c "import sys; print('Python %d.%d.%d (%d bit)' % (sys.version_info[0], sys.version_info[1], sys.version_info[2], 64 if sys.maxsize > 2**32 else 32))"
    if ($LASTEXITCODE -ne 0) { throw "无法运行解释器：$Python" }

    $versionOk = & $Python -c "import sys; print(1 if sys.version_info[:2] == (3, 4) else 0)"
    if ($versionOk.Trim() -ne "1") {
        Write-Warning "当前解释器不是 Python 3.4；XP 包必须用 3.4.10（32 位）构建，否则启动器在 XP 上跑不起来。"
        if (-not $SkipCheck) { throw "解释器版本不符（可用 -SkipCheck 强行继续，仅用于排错）" }
    }

    Write-Host "== 2/4 检查 PyInstaller ==" -ForegroundColor Cyan
    & $Python -m PyInstaller --version
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller 不可用。请执行：& `"$Python`" -m pip install `"pyinstaller==3.3.1`""
    }

    Write-Host "== 3/4 语法下限检查 + 自检 ==" -ForegroundColor Cyan
    & $Python "tools\compat_check.py" "ort_xp"
    & $Python "tools\check_env.py" --skip-source-check
    Write-Host "（自检失败不阻断打包，但请在 XP 上复跑确认）" -ForegroundColor Yellow

    Write-Host "== 4/4 打包（onedir）==" -ForegroundColor Cyan
    & $Python -m PyInstaller "packaging\ort_xp.spec" --noconfirm --clean --distpath dist --workpath build
    if ($LASTEXITCODE -ne 0) { throw "打包失败" }

    $out = Join-Path $root "dist\ORT-XP"
    Write-Host ""
    Write-Host "完成：$out" -ForegroundColor Green
    Write-Host "部署：把整个 ORT-XP 目录拷到 XP 机器，运行 ORT-XP.exe（无需安装 Python）。" -ForegroundColor Green
}
finally {
    Pop-Location
}

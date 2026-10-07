# ORT 实验室管理系统 · XP 精简客户端 —— 开发机启动脚本
#
# 用法：
#   .\tools\run_dev.ps1                                      # 用默认解析的数据目录
#   .\tools\run_dev.ps1 -DataFolder "D:\...\Data"            # 指定主程序数据目录
#   .\tools\run_dev.ps1 -Selftest                            # 只跑无界面自检
#   .\tools\run_dev.ps1 -Python "D:\Python34-32\python.exe"  # 指定解释器

param(
    [string]$DataFolder = "",
    [string]$Python = "python",
    [switch]$Selftest,
    [switch]$CheckEmail
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Push-Location $root
try {
    $arguments = @("-m", "ort_xp")
    if ($DataFolder) { $arguments += @("--data-folder", $DataFolder) }
    if ($Selftest) { $arguments += "--selftest" }
    if ($CheckEmail) { $arguments += "--check-plans-email" }

    Write-Host "启动：$Python $($arguments -join ' ')" -ForegroundColor Cyan
    & $Python @arguments
    exit $LASTEXITCODE
}
finally {
    Pop-Location
}

# ORT 实验室管理系统 · XP 精简客户端 —— Python 3.4（打包用解释器）安装脚本
#
# 为什么是 3.4.4 而不是 3.4.10：
#   python.org 上 3.4.5 之后的版本**只有源码包**，最后一个带 Windows 安装包的 3.4 是 3.4.4。
#   XP 上可用的最后一个 CPython 系列就是 3.4（3.5 起改用 VS2015 CRT，不支持 XP）。
#
# 为什么用 msiexec /a：
#   管理安装（administrative install）只把文件解包到目标目录，**不写注册表、不改系统环境**，
#   因此不需要管理员权限、也不会打乱本机其它 Python。打包只需要一个能跑的 python.exe。
#
# 用法：
#   .\tools\install_py34.ps1                     # 默认装到 D:\Python34-32
#   .\tools\install_py34.ps1 -TargetDir D:\Py34
#   .\tools\install_py34.ps1 -SkipPyInstaller    # 只装解释器
#
# 完成后用 D:\Python34-32\python.exe 作为 tools\build_xp.ps1 -Python 的参数。

param(
    [string]$TargetDir = "D:\Python34-32",
    [string]$MsiPath = "$env:TEMP\python-3.4.4.msi",
    [string]$Version = "3.4.4",
    [switch]$SkipPyInstaller
)

$ErrorActionPreference = "Stop"
$msiUrl = "https://www.python.org/ftp/python/$Version/python-$Version.msi"

function Write-Step($text) { Write-Host "== $text ==" -ForegroundColor Cyan }

Write-Step "1/5 准备安装包"
if (-not (Test-Path $MsiPath)) {
    Write-Host "下载 $msiUrl（python.org 在国内可能较慢）"
    $python = (Get-Command python).Source
    & $python -c "import urllib.request,sys; urllib.request.urlretrieve(sys.argv[1], sys.argv[2]); print('downloaded')" $msiUrl $MsiPath
    if (-not (Test-Path $MsiPath)) { throw "下载失败：$msiUrl" }
}
Write-Host ("安装包：{0}（{1:N1} MB）" -f $MsiPath, ((Get-Item $MsiPath).Length / 1MB))

Write-Step "2/5 解包到 $TargetDir（管理安装，不改系统）"
if (Test-Path $TargetDir) {
    Write-Host "目标目录已存在，先清空：$TargetDir"
    Remove-Item $TargetDir -Recurse -Force
}
New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
$log = Join-Path $env:TEMP "py34_msi_admin.log"
# ADDLOCAL=ALL：默认特性集里不含 CRT，这里把可选特性一起解出来
$msiArgs = @("/a", "`"$MsiPath`"", "/qn", "/norestart", "/l*v", "`"$log`"", "TARGETDIR=`"$TargetDir`"", "ADDLOCAL=ALL")
$proc = Start-Process msiexec.exe -ArgumentList $msiArgs -Wait -PassThru
if ($proc.ExitCode -ne 0) { throw "msiexec 退出码 $($proc.ExitCode)，日志：$log" }

Write-Step "2.5/5 补齐 32 位 VC++2010 运行库（msvcr100.dll）"
# 坑：MSI 里的 CRT 是 merge module（SharedCRT / PrivateCRT 特性），**管理安装不会解出来**，
# 于是 python.exe 直接以 0xC0000135（DLL not found）退出。这里从本机已有的副本里找一份 32 位、
# 且带 Microsoft 有效签名的 msvcr100.dll 补进目标目录。
$crtTarget = Join-Path $TargetDir "msvcr100.dll"
if (Test-Path $crtTarget) {
    Write-Host "解包已带 msvcr100.dll"
}
else {
    $candidates = @()
    foreach ($root in "C:\Program Files (x86)", "D:\Program Files (x86)", "C:\Program Files", "D:\Program Files") {
        if (Test-Path $root) {
            $candidates += Get-ChildItem $root -Filter msvcr100.dll -Recurse -Depth 4 -ErrorAction SilentlyContinue
        }
    }
    $sysWow = "C:\Windows\SysWOW64\msvcr100.dll"
    if (Test-Path $sysWow) { $candidates += Get-Item $sysWow }
    $picked = $null
    foreach ($item in $candidates) {
        $bytes = [System.IO.File]::ReadAllBytes($item.FullName)
        if ($bytes.Length -lt 0x40) { continue }
        $pe = [BitConverter]::ToUInt32($bytes, 0x3C)
        if ([BitConverter]::ToUInt16($bytes, $pe + 4) -ne 0x14C) { continue }   # 只要 x86
        $sig = Get-AuthenticodeSignature $item.FullName
        if ($sig.Status -ne "Valid" -or $sig.SignerCertificate.Subject -notlike "*Microsoft*") { continue }
        $picked = $item
        break
    }
    if (-not $picked) {
        throw ("找不到可用的 32 位 msvcr100.dll。任选一种解决办法：`n" +
               "  1) 安装 Microsoft Visual C++ 2010 SP1 Redistributable (x86)：https://www.microsoft.com/download/details.aspx?id=26999`n" +
               "  2) 从任意已装 VC++2010 的 32 位程序目录里复制 msvcr100.dll 到 $TargetDir`n" +
               "  3) 手工把本机别处的 x86\msvcr100.dll 复制到 $TargetDir")
    }
    Copy-Item $picked.FullName $crtTarget -Force
    Write-Host "已补齐 msvcr100.dll：$($picked.FullName)"
}

Write-Step "3/5 校验解释器可运行"
$pythonExe = Join-Path $TargetDir "python.exe"
if (-not (Test-Path $pythonExe)) {
    $found = Get-ChildItem $TargetDir -Filter python.exe -Recurse -File | Select-Object -First 1
    if ($found) { $pythonExe = $found.FullName }
}
if (-not (Test-Path $pythonExe)) { throw "解包后没找到 python.exe，请查看 $TargetDir 的目录结构" }
Write-Host "python.exe：$pythonExe"
& $pythonExe -c "import sys; print('Python %d.%d.%d (%d bit)' % (sys.version_info[0], sys.version_info[1], sys.version_info[2], 64 if sys.maxsize > 2**32 else 32))"
if ($LASTEXITCODE -ne 0) { throw "python.exe 无法运行（退出码 $LASTEXITCODE，通常是缺 32 位 msvcr100.dll）" }
& $pythonExe -c "import sqlite3, ssl, tkinter, smtplib; print('stdlib ok:', sqlite3.sqlite_version, ssl.OPENSSL_VERSION)"

if ($SkipPyInstaller) { Write-Host "已按参数跳过 PyInstaller 安装"; exit 0 }

Write-Step "4/5 准备 pip（ensurepip）"
& $pythonExe -m ensurepip
& $pythonExe -m pip --version

Write-Step "5/5 安装 PyInstaller 3.3.1（XP 可用的启动器）"
# 3.4 的 ssl 不读 Windows 证书库，PyPI 会报证书错误 → 用 --trusted-host 跳过校验（内网打包机可接受）
$pipHosts = @("--trusted-host", "pypi.org", "--trusted-host", "files.pythonhosted.org")
& $pythonExe -m pip install --upgrade "pip<19.2" @pipHosts
& $pythonExe -m pip install "setuptools<44.2" @pipHosts
# 依赖必须锁 3.4 时代的版本：新版 pefile / altgraph 用了 f-string 等 3.6+ 语法，3.4 下装不上
& $pythonExe -m pip install "pefile==2017.11.5" "altgraph==0.16.1" "pywin32-ctypes==0.2.0" @pipHosts
& $pythonExe -m pip install --no-deps "pyinstaller==3.3.1" @pipHosts
& $pythonExe -m PyInstaller --version

Write-Host ""
Write-Host "完成。接下来：" -ForegroundColor Green
Write-Host "  .\tools\build_xp.ps1 -Python `"$pythonExe`"" -ForegroundColor Green

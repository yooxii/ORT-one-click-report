# ORT实验室管理系统 发布打包脚本
#
# 用途：把 Release 里程序运行必需的文件（含 Data 目录下的本机设置文件）打成发布包，
#       保留目录结构，包内顶层目录 = 程序名_版本号，输出 zip 到 bin 目录。
#
# 用法：  powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\make-release.ps1
#         只打包、不重新生成： -SkipBuild
#         指定输出目录：       -OutDir D:\发布
#
# 命名：  bin\ORT实验室管理系统_0.5.zip
#        包内顶层目录同名为 ORT实验室管理系统_0.5\，解压后直接得到可运行目录。
#
# 版本号取自 Properties\AssemblyInfo.cs 的 AssemblyVersion（主.次），与「更新日志.md」的
# 「## 主.次」小节、主界面显示的版本号保持一致；发布前先把版本号与更新日志改好。

param(
    [switch]$SkipBuild,
    [string]$OutDir = (Join-Path $PSScriptRoot '..\bin')
)

$ErrorActionPreference = 'Stop'

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$releaseDir = Join-Path $repoRoot 'bin\Release'
$assemblyInfo = Join-Path $repoRoot 'Properties\AssemblyInfo.cs'

# 运行时不生成、只由生成过程产出的文件，发布包里不要
$excludeExtensions = @('.pdb', '.xml')
$excludeDirectoryNames = @('logs')

function Write-Step($text) { Write-Host "==> $text" -ForegroundColor Cyan }
function Write-Warn2($text) { Write-Host "    警告：$text" -ForegroundColor Yellow }

# 1. 版本号
# 注意：AssemblyInfo 里有一行被注释掉的示例 [assembly: AssemblyVersion("1.0.*")]，
# 因此必须匹配行首的正式那一行，并且不接受通配符
$verMatch = Select-String -Path $assemblyInfo -Pattern '^\s*\[assembly:\s*AssemblyVersion\("(\d+)\.(\d+)(?:\.\d+)*"\)' | Select-Object -First 1
if (-not $verMatch) { throw "无法从 $assemblyInfo 读到 AssemblyVersion" }
$version = "$($verMatch.Matches[0].Groups[1].Value).$($verMatch.Matches[0].Groups[2].Value)"
$baseName = "ORT实验室管理系统_$version"
Write-Step "发布版本 $version（包名 $baseName）"

# 2. 生成 Release
if ($SkipBuild) {
    Write-Step '跳过生成，直接使用现有 bin\Release'
} else {
    $msbuild = $null
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (Test-Path $vswhere) {
        $found = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -find 'MSBuild\**\Bin\MSBuild.exe' 2>$null
        if ($found) { $msbuild = @($found)[0] }
    }
    if (-not $msbuild) {
        $cmd = Get-Command MSBuild.exe -ErrorAction SilentlyContinue
        if ($cmd) { $msbuild = $cmd.Source }
    }
    if (-not $msbuild) { throw '未找到 MSBuild.exe，请安装 Visual Studio（含 .NET 桌面开发）或先把 MSBuild 加入 PATH' }

    Write-Step "重新生成 Release：$msbuild"
    & $msbuild (Join-Path $repoRoot 'ORT一键报告.csproj') `
        /p:Configuration=Release /p:SignManifests=false /p:GenerateManifests=false /v:minimal /nologo /t:Rebuild
    if ($LASTEXITCODE -ne 0) { throw "生成失败（MSBuild 退出码 $LASTEXITCODE）" }
}

if (-not (Test-Path (Join-Path $releaseDir 'ORT实验室管理系统.exe'))) {
    throw "未找到 $releaseDir\ORT实验室管理系统.exe，请先生成 Release"
}

# 3. 收集要打包的文件（保留相对路径）
Write-Step '收集程序文件'
$files = [System.Collections.Generic.List[string]]::new()
$skipped = [System.Collections.Generic.List[string]]::new()

$topFiles = Get-ChildItem $releaseDir -File | Where-Object { $excludeExtensions -notcontains $_.Extension.ToLower() }
foreach ($f in $topFiles) { $files.Add($f.Name) }

foreach ($dir in Get-ChildItem $releaseDir -Directory) {
    if ($excludeDirectoryNames -contains $dir.Name) {
        $skipped.Add("$($dir.Name)/（运行日志）")
        continue
    }
    foreach ($f in Get-ChildItem $dir.FullName -Recurse -File) {
        if ($excludeExtensions -contains $f.Extension.ToLower()) { continue }
        $files.Add($f.FullName.Substring($releaseDir.Length + 1))
    }
}

# 4. 暂存目录（保留目录结构）
$stamp = Get-Date -Format 'HHmmss'
$stageRoot = Join-Path $releaseDir "_pack_$stamp"
$stageDir = Join-Path $stageRoot $baseName
if (Test-Path $stageRoot) { Remove-Item $stageRoot -Recurse -Force }
New-Item -ItemType Directory -Path $stageDir -Force | Out-Null

Write-Step "暂存到 $stageDir"
foreach ($relative in $files) {
    $source = Join-Path $releaseDir $relative
    $target = Join-Path $stageDir $relative
    $targetParent = Split-Path $target -Parent
    if (-not (Test-Path $targetParent)) { New-Item -ItemType Directory -Path $targetParent -Force | Out-Null }
    Copy-Item $source $target -Force
}

# 5. 打包
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
$zipPath = Join-Path (Resolve-Path $OutDir).Path "$baseName.zip"
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }

Write-Step "打包 $zipPath"
Compress-Archive -Path $stageDir -DestinationPath $zipPath -CompressionLevel Optimal

# 6. 收尾与自检（自检要在删除暂存目录之前做）
$stagedLocalSettings = Join-Path $stageDir 'Data\local_settings.json'
$hasLocalSettings = Test-Path $stagedLocalSettings
$stagedChangelog = Join-Path $stageDir '更新日志.md'
$hasChangelog = Test-Path $stagedChangelog

Remove-Item $stageRoot -Recurse -Force

$logHashMatch = $true
$rootChangelog = Join-Path $repoRoot '更新日志.md'
$releaseChangelog = Join-Path $releaseDir '更新日志.md'
if ((Test-Path $rootChangelog) -and (Test-Path $releaseChangelog)) {
    $logHashMatch = (Get-FileHash $rootChangelog).Hash -eq (Get-FileHash $releaseChangelog).Hash
}

$zipSizeMb = [math]::Round((Get-Item $zipPath).Length / 1MB, 1)
Write-Host ''
Write-Host "发布包：$zipPath（$zipSizeMb MB，$($files.Count) 个文件）" -ForegroundColor Green

if (-not $hasLocalSettings) {
    Write-Warn2 '发布包里没有 Data\local_settings.json（本机设置文件），请确认 bin\Release\Data 下确实存在'
}
if (-not $hasChangelog) {
    Write-Warn2 '发布包里没有「更新日志.md」，请确认已生成 Release 且根目录日志已写完'
}
if (-not $logHashMatch) {
    Write-Warn2 '输出目录里的「更新日志.md」与根目录不一致，请重新生成后再打包'
}
if ($skipped.Count -gt 0) {
    Write-Host "    已排除：$($skipped -join '、')" -ForegroundColor DarkGray
}

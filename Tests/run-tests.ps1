# ORT 一键报告 测试运行器（不依赖 vstest 的简易版本）
#
# 用途：受限环境（LoadFrom 被系统判为“网络来源”而拒绝加载程序集）或不想安装
# 测试平台时，直接运行 NUnit 测试。测试程序集与全部依赖按字节加载（Assembly.Load(bytes)），
# 不触发基于路径的安全区域检查。
#
# 用法：  powershell -NoProfile -ExecutionPolicy Bypass -File .\Tests\run-tests.ps1
# 说明：  常规环境推荐直接用 VS 测试资源管理器或 vstest.console；本脚本只负责执行与统计，
#        覆盖 SetUp/TearDown/Test/TestCase 四类特性（当前测试集用到的全部形态）。

param(
    [string]$TestDll = (Join-Path $PSScriptRoot 'bin\Debug\ORT一键报告.Tests.dll')
)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path $TestDll)) {
    Write-Host "未找到测试程序集：$TestDll（请先构建 Tests 工程）" -ForegroundColor Red
    exit 2
}

$binDir = Split-Path (Resolve-Path $TestDll)

# 依赖解析：全部从测试输出目录按字节加载，避免任何按路径的加载（含 .exe 形式的被测程序集）
$resolver = [System.ResolveEventHandler]{
    param($sender, $e)
    $simpleName = ($e.Name -split ',')[0]
    foreach ($ext in @('.dll', '.exe')) {
        $candidate = Join-Path $binDir ($simpleName + $ext)
        if (Test-Path $candidate) {
            return [System.Reflection.Assembly]::Load([System.IO.File]::ReadAllBytes($candidate))
        }
    }
    return $null
}
[System.AppDomain]::CurrentDomain.add_AssemblyResolve($resolver)

$testAsm = [System.Reflection.Assembly]::Load([System.IO.File]::ReadAllBytes((Resolve-Path $TestDll)))

# 触发 nunit.framework 加载（测试程序集引用它，首次访问类型时经由上面的解析器字节加载）
$nunitAsm = $null
foreach ($reference in $testAsm.GetReferencedAssemblies()) {
    if ($reference.Name -eq 'nunit.framework') {
        $nunitAsm = [System.Reflection.Assembly]::Load([System.IO.File]::ReadAllBytes((Join-Path $binDir 'nunit.framework.dll')))
    }
}
if (-not $nunitAsm) {
    Write-Host '未能加载 nunit.framework' -ForegroundColor Red
    exit 2
}

$fixtureAttr = $nunitAsm.GetType('NUnit.Framework.TestFixtureAttribute')
$testAttr = $nunitAsm.GetType('NUnit.Framework.TestAttribute')
$testCaseAttr = $nunitAsm.GetType('NUnit.Framework.TestCaseAttribute')
$setUpAttr = $nunitAsm.GetType('NUnit.Framework.SetUpAttribute')
$tearDownAttr = $nunitAsm.GetType('NUnit.Framework.TearDownAttribute')

$pass = 0
$fail = 0
$failures = [System.Collections.Generic.List[string]]::new()
$flags = [System.Reflection.BindingFlags]'Public,NonPublic,Instance'

function Invoke-Case($type, $setUpMethods, $tearDownMethods, $method, $caseArgs, $caseDisplay) {
    $script:total = $script:total + 1
    $instance = [System.Activator]::CreateInstance($type)
    try {
        foreach ($s in $setUpMethods) { $s.Invoke($instance, @()) | Out-Null }
        $method.Invoke($instance, $caseArgs) | Out-Null
        $script:pass = $script:pass + 1
        Write-Host ("  通过  {0}.{1}{2}" -f $type.Name, $method.Name, $caseDisplay) -ForegroundColor Green
    }
    catch {
        $inner = $_.Exception.InnerException
        if (-not $inner) { $inner = $_.Exception }
        $script:fail = $script:fail + 1
        $line = ("  失败  {0}.{1}{2} -> {3}" -f $type.Name, $method.Name, $caseDisplay, $inner.Message)
        $script:failures.Add($line)
        Write-Host $line -ForegroundColor Red
        if ($inner.StackTrace) { Write-Host ("        " + ($inner.StackTrace -split "`n")[0]) -ForegroundColor DarkGray }
    }
    finally {
        foreach ($t in $tearDownMethods) {
            try { $t.Invoke($instance, @()) | Out-Null } catch { }
        }
    }
}

$script:total = 0
foreach ($type in $testAsm.GetTypes() | Where-Object { $_.IsClass -and $_.GetCustomAttributes($fixtureAttr, $true).Count -gt 0 }) {
    $setUpMethods = @($type.GetMethods($flags) | Where-Object { $_.GetCustomAttributes($setUpAttr, $true).Count -gt 0 })
    $tearDownMethods = @($type.GetMethods($flags) | Where-Object { $_.GetCustomAttributes($tearDownAttr, $true).Count -gt 0 })
    Write-Host ("■ {0}" -f $type.Name) -ForegroundColor Cyan

    foreach ($method in $type.GetMethods($flags) | Where-Object { $_.IsPublic -and -not $_.IsStatic }) {
        $cases = @($method.GetCustomAttributes($testCaseAttr, $true))
        if ($cases.Count -gt 0) {
            foreach ($case in $cases) {
                $args = @($case.Arguments)
                $display = ''
                if ($args.Count -gt 0) {
                    $display = '(' + (($args | ForEach-Object { if ($null -eq $_) { 'null' } else { $_.ToString() } }) -join ', ') + ')'
                }
                Invoke-Case $type $setUpMethods $tearDownMethods $method $args $display
            }
        }
        elseif ($method.GetCustomAttributes($testAttr, $true).Count -gt 0) {
            Invoke-Case $type $setUpMethods $tearDownMethods $method @() ''
        }
    }
}

Write-Host ''
Write-Host ("========== 结果：共 {0} 个用例，通过 {1}，失败 {2} ==========" -f $script:total, $script:pass, $script:fail) -ForegroundColor Yellow
if ($fail -gt 0) {
    $failures | ForEach-Object { Write-Host $_ -ForegroundColor Red }
    exit 1
}
exit 0

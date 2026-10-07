using NUnit.Framework;
using ORT一键报告.Services;
using ORT一键报告.Utils;
using System;
using System.IO;

namespace ORT一键报告.Tests
{
    /// <summary>
    /// 后台运行相关测试：开机自启的命令行与远程路径判定、单实例名称、内存清理（只读工作集）。
    /// 不写注册表、不动托盘，避免测试污染真实机器。
    /// </summary>
    [TestFixture]
    public class StartupManagerTests
    {
        /* ###############################  命令行参数  ################################ */

        [TestCase(new[] { "--background" }, true)]
        [TestCase(new[] { "--BACKGROUND" }, true)]
        [TestCase(new[] { "--background", "其他" }, true)]
        [TestCase(new[] { "其他", "--background" }, true)]
        [TestCase(new[] { "--bg" }, false)]
        [TestCase(new[] { "background" }, false)]
        [TestCase(new string[0], false)]
        [TestCase(null, false)]
        public void 后台启动参数_识别正确(string[] args, bool expected)
            => Assert.That(StartupManager.StartsInBackground(args), Is.EqualTo(expected));

        [TestCase("--background", true)]
        [TestCase("\"C:\\a\\b.exe\" --background", true)]
        [TestCase("\"C:\\a\\b.exe\"", false)]
        [TestCase("", false)]
        [TestCase(null, false)]
        [TestCase("--background2", false)]
        public void 命令行是否含后台参数(string commandLine, bool expected)
            => Assert.That(StartupManager.ContainsBackgroundArgument(commandLine), Is.EqualTo(expected));

        [Test]
        public void 自启命令行_带上程序路径并按需追加后台参数()
        {
            string toBackground = StartupManager.BuildCommandLine(true);
            string foreground = StartupManager.BuildCommandLine(false);

            Assert.That(toBackground, Does.Contain(StartupManager.BackgroundArgument));
            Assert.That(toBackground, Does.StartWith("\""));
            Assert.That(foreground, Does.Not.Contain(StartupManager.BackgroundArgument));
            // 生成的自启命令行必须能被自己的参数解析认出来（下次启动才知道要进后台）
            Assert.That(StartupManager.ContainsBackgroundArgument(foreground), Is.False);
            Assert.That(StartupManager.StartsInBackground([StartupManager.BackgroundArgument]), Is.True);
        }

        [Test]
        public void 相对路径_不是远程也不含驱动器盘符()
        {
            string relative = StartupManager.RelativeExePath();

            Assert.That(relative, Is.Not.Null);
            // 相对路径形态（不含盘符）或跨盘符时退回的完整路径，两种都不该是 UNC
            Assert.That(relative.StartsWith(@"\\", StringComparison.Ordinal), Is.False);
        }

        /* ###############################  远程路径判定  ################################ */

        [Test]
        public void 远程路径_UNC与斜杠开头都算远程()
        {
            Assert.That(StartupManager.TryDescribeRemoteLocation(@"\\server\share\ORT\app.exe"), Is.EqualTo(@"\\server\share\ORT\app.exe"));
            Assert.That(StartupManager.TryDescribeRemoteLocation("//server/share/ORT/app.exe"), Is.Not.Null);
        }

        [TestCase(null)]
        [TestCase("")]
        [TestCase("   ")]
        public void 远程路径_空输入不算远程(string path)
            => Assert.That(StartupManager.TryDescribeRemoteLocation(path), Is.Null);

        [Test]
        public void 远程路径_本机固定磁盘不算远程()
        {
            string local = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "notepad.exe");

            Assert.That(StartupManager.TryDescribeRemoteLocation(local), Is.Null);
            Assert.That(StartupManager.IsRemoteLocation(local), Is.False);
        }

        /* ###############################  SUBST 盘符  ################################ */

        [TestCase("C:")]
        [TestCase("C:\\")]
        public void SUBST判定_系统盘不是虚拟盘(string root)
            => Assert.That(StartupManager.IsSubstitutedDrive(root), Is.False);

        [TestCase(null)]
        [TestCase("")]
        [TestCase("不是盘符")]
        [TestCase(@"\\server\share")]
        public void SUBST判定_非盘符输入返回false(string root)
            => Assert.That(StartupManager.IsSubstitutedDrive(root), Is.False);

        /* ###############################  单实例  ################################ */

        [Test]
        public void 单实例_互斥体与管道名同作用域且非空()
        {
            Assert.That(SingleInstanceManager.MutexName, Does.StartWith(@"Local\"));
            Assert.That(SingleInstanceManager.PipeName, Is.Not.Empty);
            Assert.That(SingleInstanceManager.ActivateMessage, Is.Not.Empty);
            // 同机不同用户允许各开一个，所以不能用 Global\ 作用域
            Assert.That(SingleInstanceManager.MutexName, Does.Not.StartWith(@"Global\"));
        }

        /* ###############################  内存清理  ################################ */

        [Test]
        public void 内存清理_可重复调用且能读到工作集()
        {
            MemoryTrimmer.Trim("单元测试");
            long first = MemoryTrimmer.LastWorkingSetKb;
            Assert.That(first, Is.GreaterThan(0));

            MemoryTrimmer.Trim("单元测试-第二次");
            Assert.That(MemoryTrimmer.LastWorkingSetKb, Is.GreaterThan(0));
            Assert.That(MemoryTrimmer.WorkingSetKb(), Is.GreaterThan(0));
        }
    }
}

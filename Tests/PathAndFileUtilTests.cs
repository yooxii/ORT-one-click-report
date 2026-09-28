using NUnit.Framework;
using ORT一键报告.Utils;
using System;
using System.IO;

namespace ORT一键报告.Tests
{
    /// <summary>
    /// 路径与文件工具测试：文件夹路径规范化（FolderUtil）、UNC 拆分（NetworkDriveMapper）、
    /// SQLite 库文件头判定（SqliteFileUtil.IsWalFile）。
    /// </summary>
    [TestFixture]
    public class PathAndFileUtilTests
    {
        /* ###############################  路径规范化  ################################ */

        [Test]
        public void 规范化_去引号与末尾分隔符()
            => Assert.That(FolderUtil.Normalize("  \"C:\\Temp\\ort\\\"  "), Is.EqualTo("C:\\Temp\\ort"));

        [Test]
        public void 规范化_环境变量展开()
        {
            string expected = Path.GetFullPath(Environment.ExpandEnvironmentVariables("%TEMP%\\ort_test"));
            Assert.That(FolderUtil.Normalize("%TEMP%\\ort_test"), Is.EqualTo(expected));
        }

        [Test]
        public void 规范化_盘符根目录保留分隔符()
            => Assert.That(FolderUtil.Normalize("C:\\"), Is.EqualTo("C:\\"));

        [TestCase(null)]
        [TestCase("")]
        [TestCase("   ")]
        [TestCase("\"\"")]
        public void 规范化_空输入返回空(string folder)
            => Assert.That(FolderUtil.Normalize(folder), Is.Null);

        [Test]
        public void 网络路径判断()
        {
            Assert.That(FolderUtil.IsNetworkPath(@"\\server\share\sub"), Is.True);
            Assert.That(FolderUtil.IsNetworkPath("//server/share"), Is.True);
            Assert.That(FolderUtil.IsNetworkPath("C:\\Windows"), Is.False);
            Assert.That(FolderUtil.IsNetworkPath(null), Is.False);
        }

        /* ###############################  UNC 拆分  ################################ */

        [Test]
        public void UNC拆分_共享根与相对路径()
        {
            bool ok = NetworkDriveMapper.TrySplitUnc(@"\\server\share\a\b", out string root, out string relative);

            Assert.That(ok, Is.True);
            Assert.That(root, Is.EqualTo(@"\\server\share"));
            Assert.That(relative, Is.EqualTo(@"a\b"));
        }

        [Test]
        public void UNC拆分_仅共享根()
        {
            bool ok = NetworkDriveMapper.TrySplitUnc(@"\\server\share", out string root, out string relative);

            Assert.That(ok, Is.True);
            Assert.That(root, Is.EqualTo(@"\\server\share"));
            Assert.That(relative, Is.Empty);
        }

        [Test]
        public void UNC拆分_末尾分隔符被规范化掉()
        {
            bool ok = NetworkDriveMapper.TrySplitUnc(@"\\server\share\dir\", out string root, out string relative);

            Assert.That(ok, Is.True);
            Assert.That(root, Is.EqualTo(@"\\server\share"));
            Assert.That(relative, Is.EqualTo("dir"));
        }

        [TestCase(@"C:\temp")]
        [TestCase(@"server\share")]
        [TestCase(@"\\server")]
        [TestCase(null)]
        public void UNC拆分_非UNC输入返回false(string path)
            => Assert.That(NetworkDriveMapper.TrySplitUnc(path, out _, out _), Is.False);

        /* ###############################  SQLite 文件头  ################################ */

        [Test]
        public void SQLite文件头_WAL模式判定()
        {
            byte[] header = new byte[100];
            header[18] = 2;    // 写版本 = WAL
            header[19] = 2;    // 读版本 = WAL
            string walFile = WriteTempDb(header);

            byte[] journal = new byte[100];
            journal[18] = 1;   // 回滚日志模式
            journal[19] = 1;
            string journalFile = WriteTempDb(journal);

            try
            {
                Assert.That(SqliteFileUtil.IsWalFile(walFile), Is.True);
                Assert.That(SqliteFileUtil.IsWalFile(journalFile), Is.False);
            }
            finally
            {
                SafeDelete(walFile);
                SafeDelete(journalFile);
            }
        }

        [Test]
        public void SQLite文件头_文件过短或不存在_返回false()
        {
            string shortFile = Path.Combine(Path.GetTempPath(), $"ort_db_short_{Guid.NewGuid():N}.db");
            File.WriteAllBytes(shortFile, new byte[10]);
            try
            {
                Assert.That(SqliteFileUtil.IsWalFile(shortFile), Is.False);   // 读不满 20 字节
                Assert.That(SqliteFileUtil.IsWalFile(Path.Combine(Path.GetTempPath(), $"ort_missing_{Guid.NewGuid():N}.db")), Is.False);
                Assert.That(SqliteFileUtil.IsWalFile(null), Is.False);
            }
            finally
            {
                SafeDelete(shortFile);
            }
        }

        private static string WriteTempDb(byte[] content)
        {
            string path = Path.Combine(Path.GetTempPath(), $"ort_db_test_{Guid.NewGuid():N}.db");
            File.WriteAllBytes(path, content);
            return path;
        }

        private static void SafeDelete(string path)
        {
            try { File.Delete(path); } catch { /* 测试清理失败不阻塞 */ }
        }
    }
}

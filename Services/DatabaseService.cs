using FreeSql;
using NLog;
using ORT一键报告.Models;
using System;
using System.IO;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 全局数据库服务（SQLite + FreeSql）。
    /// 数据库文件保存在程序根目录的 Data 文件夹中；
    /// 嵌入对象（OLE）等附件保存在 Data\OleFiles 文件夹中。
    /// 考虑最大并发 10 人次以内：启用 WAL 模式，连接池上限 10。
    /// </summary>
    public class DatabaseService : IDisposable
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// 数据根目录（程序根目录\Data）
        /// </summary>
        public string DataDir { get; }

        /// <summary>
        /// 数据库文件完整路径
        /// </summary>
        public string DbPath { get; }

        /// <summary>
        /// 附件（OLE提取文件/上传的SN文件）目录
        /// </summary>
        public string OleDir { get; }

        /// <summary>
        /// FreeSql 实例
        /// </summary>
        public IFreeSql FreeSql { get; }

        public DatabaseService()
        {
            DataDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");
            OleDir = Path.Combine(DataDir, "OleFiles");
            Directory.CreateDirectory(DataDir);
            Directory.CreateDirectory(OleDir);
            DbPath = Path.Combine(DataDir, "ort_plans.db");

            string connStr = $"Data Source={DbPath};Pooling=true;Min Pool Size=1;Max Pool Size=10";
            FreeSql = new FreeSqlBuilder()
                .UseConnectionString(DataType.Sqlite, connStr)
                .UseAutoSyncStructure(true) // 首次运行自动建表/同步结构
                .Build();

            // WAL 模式提升并发读写性能（最大并发10人次以内足够）
            FreeSql.Ado.ExecuteNonQuery("PRAGMA journal_mode=WAL;");
            _logger.Info($"数据库初始化完成: {DbPath}");
        }

        /// <summary>
        /// 将相对 OleDir 的附件路径转为绝对路径；绝对路径原样返回
        /// </summary>
        public string ResolveAttachmentPath(string relativeOrAbsolute)
        {
            if (string.IsNullOrWhiteSpace(relativeOrAbsolute))
            {
                return null;
            }
            return Path.IsPathRooted(relativeOrAbsolute)
                ? relativeOrAbsolute
                : Path.Combine(OleDir, relativeOrAbsolute);
        }

        public void Dispose()
        {
            FreeSql?.Dispose();
            _logger.Info("数据库连接已释放");
        }
    }
}

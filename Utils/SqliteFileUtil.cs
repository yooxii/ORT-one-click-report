using NLog;
using System;
using System.Data.SQLite;
using System.IO;

namespace ORT一键报告.Utils
{
    /// <summary>
    /// SQLite 库文件工具：
    /// 1) 读文件头判断库是不是 WAL（预写日志）模式（头偏移 18/19 都是 2）；
    /// 2) 把 WAL 模式的库转成回滚日志模式（rollback journal）。
    ///
    /// 为什么需要：SQLite 的 WAL 依赖 -shm 文件的共享内存映射，**在网络共享（SMB/UNC）上不可用**。
    /// 库文件一旦是 WAL 模式，放到 \\服务器\共享 上就会直接打不开，
    /// System.Data.SQLite 报的正是 "unable to open database file"；
    /// 而且这时连 PRAGMA journal_mode=DELETE 都执行不了（打开就失败），
    /// 所以只能在**本地磁盘**上转换：先复制到临时目录转换，再写回目标路径。
    /// </summary>
    public static class SqliteFileUtil
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        /// <summary>SQLite 文件头：读写版本字节的位置（1=回滚日志，2=WAL）</summary>
        private const int FormatVersionOffset = 18;

        /// <summary>
        /// 库文件是否是 WAL 模式（文件头偏移 18/19 都是 2）；不存在或读不到返回 false
        /// </summary>
        public static bool IsWalFile(string dbPath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(dbPath) || !File.Exists(dbPath))
                {
                    return false;
                }
                byte[] header = new byte[20];
                using (FileStream stream = new(dbPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    int read = 0;
                    while (read < header.Length)
                    {
                        int n = stream.Read(header, read, header.Length - read);
                        if (n <= 0)
                        {
                            break;
                        }
                        read += n;
                    }
                    if (read < 20)
                    {
                        return false;
                    }
                }
                return header[FormatVersionOffset] == 2 && header[FormatVersionOffset + 1] == 2;
            }
            catch (Exception ex)
            {
                Logger.Warn($"读取 SQLite 文件头失败（{dbPath}）: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 生成 sourceDb 的"回滚日志模式"副本写到 target（target 可以在网络共享上）；
        /// 已提交但还在 -wal 里的事务会被一并并回主库，不会丢数据。
        /// </summary>
        public static bool CreateRollbackCopy(string sourceDb, string target, out string error)
        {
            error = null;
            if (string.IsNullOrWhiteSpace(sourceDb) || !File.Exists(sourceDb))
            {
                error = "源数据库文件不存在";
                return false;
            }
            string tempDir = Path.Combine(Path.GetTempPath(), "ort_sqlite_" + Guid.NewGuid().ToString("N"));
            try
            {
                Directory.CreateDirectory(tempDir);
                string tempDb = Path.Combine(tempDir, Path.GetFileName(sourceDb));
                File.Copy(sourceDb, tempDb, overwrite: true);
                CopyIfExists(sourceDb + "-wal", tempDb + "-wal");
                CopyIfExists(sourceDb + "-shm", tempDb + "-shm");
                if (!ConvertToRollback(tempDb, out error))
                {
                    return false;
                }
                string targetDir = Path.GetDirectoryName(target);
                if (!string.IsNullOrEmpty(targetDir))
                {
                    Directory.CreateDirectory(targetDir);
                }
                File.Copy(tempDb, target, overwrite: true);
                // 主库里已经并入了 WAL 内容，目标旁边残留的 -wal/-shm 不再需要
                DeleteIfExists(target + "-wal");
                DeleteIfExists(target + "-shm");
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
            finally
            {
                TryDeleteDirectory(tempDir);
            }
        }

        /// <summary>
        /// 就地转换：把 dbPath 的库从 WAL 模式转成回滚日志模式（中间经过本地临时目录）
        /// </summary>
        public static bool TryConvertToRollbackInPlace(string dbPath, out string error)
            => CreateRollbackCopy(dbPath, dbPath, out error);

        /// <summary>
        /// 在本地磁盘上把库转成回滚日志模式（PRAGMA journal_mode=DELETE，WAL 内容自动并回主库）
        /// </summary>
        private static bool ConvertToRollback(string localDbPath, out string error)
        {
            error = null;
            try
            {
                SQLiteConnectionStringBuilder builder = new()
                {
                    DataSource = localDbPath,
                    Pooling = false,
                    DefaultTimeout = 30,
                    BusyTimeout = 30000
                };
                using SQLiteConnection connection = new(builder.ConnectionString);
                connection.Open();
                // 先把 WAL 内容并回主库（非 WAL 库是无害的空操作）
                try
                {
                    using SQLiteCommand checkpoint = new("PRAGMA wal_checkpoint(TRUNCATE);", connection);
                    checkpoint.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    Logger.Warn($"wal_checkpoint 跳过: {ex.Message}");
                }
                using (SQLiteCommand command = new("PRAGMA journal_mode=DELETE;", connection))
                {
                    object mode = command.ExecuteScalar();
                    Logger.Info($"SQLite 日志模式已转换为: {mode}");
                }
                connection.Close();
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        private static void CopyIfExists(string from, string to)
        {
            if (File.Exists(from))
            {
                File.Copy(from, to, overwrite: true);
            }
        }

        private static void DeleteIfExists(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn($"删除文件失败（{path}）: {ex.Message}");
            }
        }

        private static void TryDeleteDirectory(string dir)
        {
            try
            {
                if (Directory.Exists(dir))
                {
                    Directory.Delete(dir, recursive: true);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn($"清理临时目录失败（{dir}）: {ex.Message}");
            }
        }
    }
}

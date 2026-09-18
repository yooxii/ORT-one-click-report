using FreeSql;
using NLog;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.IO;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 数据文件夹迁移：把当前数据文件夹里的业务数据（数据库 / 附件 OleFiles / 配图 PlanImages）
    /// 复制到新的数据文件夹——换到远程共享时沿用现有数据。
    /// 本机设置文件（local_settings.json）等本机内容不复制。
    /// </summary>
    public static class DataFolderMigrator
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        /// <summary>数据文件夹里的数据项（相对路径）</summary>
        private static readonly string[] DataItems = ["OleFiles", "PlanImages"];

        /// <summary>
        /// 把 sourceDir 里的数据复制到 targetDir（目标已存在的同名文件不覆盖，避免覆盖共享文件夹里别人正在用的数据）
        /// </summary>
        /// <returns>(复制的文件数, 错误列表)</returns>
        public static (int Copied, List<string> Errors) CopyTo(string sourceDir, string targetDir, IFreeSql freeSql)
        {
            List<string> errors = [];
            int copied = 0;
            string source = FolderUtil.Normalize(sourceDir);
            string target = FolderUtil.Normalize(targetDir);
            if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(target)
                || string.Equals(source, target, StringComparison.OrdinalIgnoreCase) || !Directory.Exists(source))
            {
                return (0, errors);
            }
            bool targetIsNetwork = FolderUtil.IsNetworkPath(target);
            // 数据库还在使用中：先把 WAL 内容并回主库，避免复制到"缺最新事务"的库
            if (freeSql != null)
            {
                try
                {
                    freeSql.Ado.ExecuteNonQuery("PRAGMA wal_checkpoint(TRUNCATE);");
                }
                catch (Exception ex)
                {
                    Logger.Warn($"复制数据前的 wal_checkpoint 失败（继续复制）: {ex.Message}");
                }
            }
            try
            {
                Directory.CreateDirectory(target);
            }
            catch (Exception ex)
            {
                errors.Add($"无法创建目标文件夹「{target}」：{ex.Message}");
                return (copied, errors);
            }

            // 数据库：目标是网络共享时不能直接拷 WAL 模式的库（拷过去打不开），
            // 先在本地临时目录转成回滚日志模式再写过去
            string dbFrom = Path.Combine(source, "ort_plans.db");
            string dbTo = Path.Combine(target, "ort_plans.db");
            if (File.Exists(dbFrom) && !File.Exists(dbTo))
            {
                try
                {
                    if (targetIsNetwork)
                    {
                        if (SqliteFileUtil.CreateRollbackCopy(dbFrom, dbTo, out string dbError))
                        {
                            copied++;
                        }
                        else
                        {
                            errors.Add($"ort_plans.db: {dbError}");
                        }
                    }
                    else
                    {
                        File.Copy(dbFrom, dbTo, overwrite: false);
                        copied++;
                    }
                }
                catch (Exception ex)
                {
                    errors.Add($"ort_plans.db: {ex.Message}");
                }
            }

            foreach (string item in DataItems)
            {
                string from = Path.Combine(source, item);
                string to = Path.Combine(target, item);
                try
                {
                    if (Directory.Exists(from))
                    {
                        copied += CopyDirectory(from, to, errors);
                    }
                }
                catch (Exception ex)
                {
                    errors.Add($"{item}: {ex.Message}");
                }
            }
            Logger.Info($"数据文件夹迁移完成：{source} → {target}（{(targetIsNetwork ? "网络共享" : "本地")}），复制 {copied} 个文件，失败 {errors.Count} 项");
            return (copied, errors);
        }

        private static int CopyDirectory(string source, string target, List<string> errors)
        {
            int copied = 0;
            Directory.CreateDirectory(target);
            foreach (string file in Directory.GetFiles(source))
            {
                string name = Path.GetFileName(file);
                if (name.StartsWith("~$", StringComparison.Ordinal))
                {
                    continue;
                }
                string to = Path.Combine(target, name);
                if (File.Exists(to))
                {
                    continue;
                }
                try
                {
                    File.Copy(file, to, overwrite: false);
                    copied++;
                }
                catch (Exception ex)
                {
                    errors.Add($"{name}: {ex.Message}");
                }
            }
            foreach (string dir in Directory.GetDirectories(source))
            {
                try
                {
                    copied += CopyDirectory(dir, Path.Combine(target, Path.GetFileName(dir)), errors);
                }
                catch (Exception ex)
                {
                    errors.Add($"{Path.GetFileName(dir)}: {ex.Message}");
                }
            }
            return copied;
        }
    }
}

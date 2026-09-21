using FreeSql;
using NLog;
using ORT一键报告.Models;
using ORT一键报告.Utils;
using System;
using System.IO;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 全局数据库服务（SQLite + FreeSql）。
    /// 数据库文件、附件（OleFiles）、配图（PlanImages）统一放在**数据文件夹**里
    /// （见「设置 → 数据文件夹」；可以是本地目录，也可以是 \\服务器\共享 这类远程文件夹，
    /// 多台电脑指向同一个文件夹即可共用同一套数据）。
    /// 考虑最大并发 10 人次以内：本地目录用 WAL + 连接池；网络文件夹用回滚日志 +
    /// 较长的忙等待（SQLite 的 WAL 依赖共享内存，在 SMB 共享上不可用）。
    /// </summary>
    public class DatabaseService : IDisposable
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// 设置里配置的数据文件夹（可能是 UNC 路径；实际使用的见 <see cref="DataDir"/>）
        /// </summary>
        public string ConfiguredDataFolder { get; private set; }

        /// <summary>
        /// 数据根目录（实际使用：UNC 路径会被映射成网络驱动器后的盘符路径）
        /// </summary>
        public string DataDir { get; private set; }

        /// <summary>
        /// 数据库文件完整路径
        /// </summary>
        public string DbPath { get; private set; }

        /// <summary>
        /// 附件（OLE提取文件/上传的SN文件）目录
        /// </summary>
        public string OleDir { get; private set; }

        /// <summary>
        /// 计划索引抽出的测试项目配图目录（随数据库一起走，多客户端共用）
        /// </summary>
        public string PlanImagesDir { get; private set; }

        /// <summary>
        /// 数据文件夹是否在网络上（UNC 或映射的网络驱动器）
        /// </summary>
        public bool IsNetworkFolder { get; private set; }

        /// <summary>
        /// FreeSql 实例
        /// </summary>
        public IFreeSql FreeSql { get; private set; }

        /// <summary>
        /// 是否已回退到程序目录下的默认 Data 文件夹（配置路径启动失败时才会为 true，仅本次运行生效）
        /// </summary>
        public bool IsUsingFallbackDataFolder { get; private set; }

        /// <summary>
        /// 回退到默认数据文件夹的原因（即配置路径初始化失败时抛出的错误信息，未回退时为 null）
        /// </summary>
        public string FallbackReason { get; private set; }

        /// <summary>
        /// 程序目录下的默认数据文件夹（本机设置未配置或配置路径启动失败时使用）
        /// </summary>
        public static string DefaultDataFolder => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");

        public DatabaseService()
        {
            // 数据文件夹可在设置中修改（保存在程序目录的本机设置文件里，重启生效）
            string configured = AppSettingsService.ResolveDataFolder();
            ConfiguredDataFolder = configured;

            try
            {
                InitializeDataFolder(configured);
            }
            catch (DataFolderUnavailableException ex)
            {
                // 配置路径启动失败：若配置值就是默认 Data，直接抛（避免死循环重试）；
                // 否则释放首次尝试可能已建立的 FreeSql 连接，改用程序目录下的默认 Data 重试一次。
                // 回退只对本次运行生效，不写回 local_settings.json——网络共享临时不可达时不应把用户配置清掉。
                DisposeFreeSqlQuietly();
                string normalizedConfigured = FolderUtil.Normalize(configured);
                string normalizedDefault = FolderUtil.Normalize(DefaultDataFolder);
                if (string.Equals(normalizedConfigured, normalizedDefault, StringComparison.OrdinalIgnoreCase))
                {
                    throw;
                }
                _logger.Warn(ex, $"数据文件夹「{configured}」启动失败，回退到默认数据文件夹「{DefaultDataFolder}」");
                IsUsingFallbackDataFolder = true;
                FallbackReason = ex.Message;
                InitializeDataFolder(DefaultDataFolder);
            }
        }

        /// <summary>
        /// 按指定文件夹初始化数据目录与 FreeSql 实例；失败时抛出 <see cref="DataFolderUnavailableException"/>。
        /// </summary>
        private void InitializeDataFolder(string folder)
        {
            // SQLite 打不开 UNC 路径的库文件（unable to open database file），
            // 所以填 UNC 时先自动映射成网络驱动器，再用盘符路径访问
            if (NetworkDriveMapper.IsUncPath(folder))
            {
                if (!NetworkDriveMapper.TryMap(folder, out string mappedFolder, out string mapError))
                {
                    throw new DataFolderUnavailableException(
                        $"数据文件夹是网络路径（{folder}），SQLite 不能直接使用 UNC 路径，" +
                        $"自动映射为网络驱动器也失败了：{mapError}{Environment.NewLine}{Environment.NewLine}" +
                        "请先手工把共享映射成盘符，例如在命令行执行：" + Environment.NewLine +
                        "    net use Z: \\\\服务器\\共享 /persistent:yes" + Environment.NewLine +
                        "然后在「设置 → 数据文件夹」里改成 Z:\\对应目录（或直接改程序目录下 Data\\local_settings.json 的 DataFolder）后重启。");
                }
                DataDir = mappedFolder;
            }
            else
            {
                DataDir = folder;
            }
            if (!FolderUtil.TryPrepare(DataDir, out string error))
            {
                throw new DataFolderUnavailableException(
                    $"数据文件夹不可用：{error}{Environment.NewLine}{Environment.NewLine}" +
                    "请检查该文件夹是否存在、是否有读写权限（网络共享请确认能访问），" +
                    $"或修改程序目录下 Data\\local_settings.json 里的 DataFolder（当前值：{folder}）。");
            }
            IsNetworkFolder = FolderUtil.IsNetworkPath(DataDir);
            DbPath = Path.Combine(DataDir, "ort_plans.db");
            OleDir = Path.Combine(DataDir, "OleFiles");
            PlanImagesDir = Path.Combine(DataDir, "PlanImages");
            Directory.CreateDirectory(OleDir);
            Directory.CreateDirectory(PlanImagesDir);

            // 网络共享上的库不能是 WAL 模式（SQLite 的 WAL 依赖共享内存，SMB 上打不开），
            // 打开之前先把它转成回滚日志模式；否则会一路报 unable to open database file
            if (IsNetworkFolder)
            {
                EnsureRollbackJournalForNetworkFolder();
            }

            string connStr = $"Data Source={DbPath};Pooling=true;Min Pool Size=1;Max Pool Size=10;Default Timeout=30";
            FreeSql = new FreeSqlBuilder()
                .UseConnectionString(DataType.Sqlite, connStr)
                .UseAutoSyncStructure(true) // 首次运行自动建表/同步结构
                .Build();

            // 确认真的打得开：打不开就明确报错（否则界面能起来但每次操作都失败，只留下 "unable to open database file"）
            VerifyOpenable();
            ConfigureJournal();
            MigrateLegacyPlansToRequisitions();
            _logger.Info($"数据库初始化完成: {DbPath}（数据文件夹: {DataDir}{(IsNetworkFolder ? "，网络共享" : "")}，日志模式: {CurrentJournalMode() ?? "未知"}{(IsUsingFallbackDataFolder ? "，已回退到默认目录" : "")}）");
        }

        /// <summary>
        /// 释放首次尝试建立的 FreeSql 连接（回退前调用，避免遗留未释放的连接池）
        /// </summary>
        private void DisposeFreeSqlQuietly()
        {
            if (FreeSql == null)
            {
                return;
            }
            try
            {
                FreeSql.Dispose();
            }
            catch (Exception ex)
            {
                _logger.Warn($"释放回退前的数据库连接失败: {ex.Message}");
            }
            finally
            {
                FreeSql = null;
            }
        }

        /// <summary>
        /// 网络共享上的库若是 WAL 模式，就地转成回滚日志模式（本地临时目录转换后写回）
        /// </summary>
        private void EnsureRollbackJournalForNetworkFolder()
        {
            try
            {
                if (!File.Exists(DbPath) || !SqliteFileUtil.IsWalFile(DbPath))
                {
                    return;
                }
                _logger.Warn($"数据库处于 WAL 模式，网络共享上无法打开，正在转换为回滚日志模式: {DbPath}");
                if (SqliteFileUtil.TryConvertToRollbackInPlace(DbPath, out string convertError))
                {
                    _logger.Info("已转换为回滚日志模式（DELETE）");
                }
                else
                {
                    _logger.Error($"转换为回滚日志模式失败: {convertError}");
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"检查/转换数据库日志模式失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 数据库打不开时给出带诊断信息的错误（路径、是否网络共享、是否 WAL、SQLite 原始错误）
        /// </summary>
        private void VerifyOpenable()
        {
            try
            {
                FreeSql.Ado.ExecuteScalar("SELECT COUNT(*) FROM sqlite_master;");
            }
            catch (Exception ex)
            {
                string wal = SqliteFileUtil.IsWalFile(DbPath) ? "是" : "否";
                throw new DataFolderUnavailableException(
                    $"数据文件夹里的数据库打不开：{ex.Message}{Environment.NewLine}{Environment.NewLine}" +
                    $"数据库文件：{DbPath}{Environment.NewLine}" +
                    $"网络共享：{(IsNetworkFolder ? "是" : "否")}；WAL 模式：{wal}{Environment.NewLine}{Environment.NewLine}" +
                    "常见原因：库放在网络共享上且处于 WAL 模式（SQLite 的 WAL 依赖共享内存，共享文件夹上不可用）、" +
                    "文件被别的程序独占、没有读写权限，或文件损坏。可确认该文件能否打开、" +
                    "旁边是否有残留的 ort_plans.db-wal，或在设置里换回本地数据文件夹后重启。");
            }
        }

        /// <summary>
        /// 当前日志模式（诊断用）
        /// </summary>
        private string CurrentJournalMode()
        {
            try
            {
                return FreeSql.Ado.ExecuteScalar("PRAGMA journal_mode;") as string;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 日志模式与并发等待：
        /// - 本地目录：WAL（读写并发好，崩溃恢复快）；
        /// - 网络共享：SQLite 的 WAL 需要共享内存映射，在 SMB 上不可用，改用回滚日志（TRUNCATE）
        ///   并把同步级别提到 FULL；同时设置忙等待，多台电脑同时写时先排队重试而不是直接报 database is locked。
        /// </summary>
        private void ConfigureJournal()
        {
            try
            {
                if (IsNetworkFolder)
                {
                    FreeSql.Ado.ExecuteNonQuery("PRAGMA journal_mode=TRUNCATE;");
                    FreeSql.Ado.ExecuteNonQuery("PRAGMA synchronous=FULL;");
                }
                else
                {
                    FreeSql.Ado.ExecuteNonQuery("PRAGMA journal_mode=WAL;");
                }
                FreeSql.Ado.ExecuteNonQuery("PRAGMA busy_timeout=30000;");
            }
            catch (Exception ex)
            {
                _logger.Warn($"设置数据库日志模式/忙等待失败（继续运行）: {ex.Message}");
            }
        }

        /// <summary>
        /// 旧版数据迁移：单表（plans）拆分为 plans/requisitions 后，将 plans 表中遗留的领退字段一次性迁移到 requisitions 表；
        /// 同时清洗旧日期文本（无法解析的置空，避免 FreeSql 读取 DateTime 崩溃）
        /// </summary>
        private void MigrateLegacyPlansToRequisitions()
        {
            try
            {
                // 清理早期版本创建的已废弃唯一索引（回线RT工令/WorkOrder 可对应多条记录）
                FreeSql.Ado.ExecuteNonQuery("DROP INDEX IF EXISTS uk_req_return_rt");
                FreeSql.Ado.ExecuteNonQuery("DROP INDEX IF EXISTS uk_req_workorder");

                // 仅当 plans 表仍存在旧领退列时执行清洗
                bool hasLegacyColumn = FreeSql.Ado.ExecuteScalar(
                    "SELECT COUNT(*) FROM pragma_table_info('plans') WHERE name = 'RequisitionNo'") is long l && l > 0;
                if (hasLegacyColumn)
                {
                    // 清洗旧日期文本：无法被 SQLite DateTime 转换解析的（非 ISO 格式）置空
                    foreach (string col in new[] { "StartDate", "EndDate", "RequisitionDate", "ReturnDate", "StockInDate", "CreatedAt", "UpdatedAt" })
                    {
                        FreeSql.Ado.ExecuteNonQuery(
                            $"UPDATE plans SET {col} = NULL WHERE {col} IS NOT NULL AND {col} != '' AND {col} NOT LIKE '____-__-__%'");
                    }
                }

                // 仅当 requisitions 表为空（或含旧日期文本需重迁）且 plans 表仍存在旧领退列时执行迁移
                long badDates = (long)FreeSql.Ado.ExecuteScalar(
                    "SELECT COUNT(*) FROM requisitions WHERE (RequisitionDate IS NOT NULL AND RequisitionDate != '' AND RequisitionDate NOT LIKE '____-__-__%') OR (ReturnDate IS NOT NULL AND ReturnDate != '' AND ReturnDate NOT LIKE '____-__-__%') OR (StockInDate IS NOT NULL AND StockInDate != '' AND StockInDate NOT LIKE '____-__-__%')");
                if (badDates > 0)
                {
                    // 旧迁移写入的日期文本无法解析，清空后重新按行解析迁移
                    FreeSql.Ado.ExecuteNonQuery("DELETE FROM requisitions");
                    _logger.Info("检测到领退表旧日期文本，已清空准备重新迁移");
                }
                bool requisitionsEmpty = FreeSql.Select<Requisition>().Count() == 0;
                if (!hasLegacyColumn || !requisitionsEmpty)
                {
                    return;
                }
                int migrated = 0;
                System.Data.DataTable rows = FreeSql.Ado.ExecuteDataTable(
                    "SELECT RequisitionDate, RequisitionNo, ModelName, OutQty, SN, SnFilePath, DC, Rev, WorkOrder, ReturnRtOrder, ReturnQty, LineNo, ReturnDate, StockInNo, StockInQty, StockInDate, Remark, CreatedBy, CreatedAt, UpdatedBy, UpdatedAt FROM plans WHERE RequisitionNo IS NOT NULL AND RequisitionNo != ''");
                foreach (System.Data.DataRow row in rows.Rows)
                {
                    Requisition req = new()
                    {
                        RequisitionDate = ParseDate(row["RequisitionDate"] as string),
                        RequisitionNo = NullIfEmpty(row["RequisitionNo"] as string),
                        ModelName = NullIfEmpty(row["ModelName"] as string),
                        OutQty = NullIfEmpty(row["OutQty"] as string),
                        SN = NullIfEmpty(row["SN"] as string),
                        SnFilePath = NullIfEmpty(row["SnFilePath"] as string),
                        DC = NullIfEmpty(row["DC"] as string),
                        Rev = NullIfEmpty(row["Rev"] as string),
                        WorkOrder = NullIfEmpty(row["WorkOrder"] as string),
                        ReturnRtOrder = NullIfEmpty(row["ReturnRtOrder"] as string),
                        ReturnQty = NullIfEmpty(row["ReturnQty"] as string),
                        LineNo = NullIfEmpty(row["LineNo"] as string),
                        ReturnDate = ParseDate(row["ReturnDate"] as string),
                        StockInNo = NullIfEmpty(row["StockInNo"] as string),
                        StockInQty = NullIfEmpty(row["StockInQty"] as string),
                        StockInDate = ParseDate(row["StockInDate"] as string),
                        Remark = NullIfEmpty(row["Remark"] as string),
                        CreatedBy = NullIfEmpty(row["CreatedBy"] as string),
                        CreatedAt = ParseDate(row["CreatedAt"] as string),
                        UpdatedBy = NullIfEmpty(row["UpdatedBy"] as string),
                        UpdatedAt = ParseDate(row["UpdatedAt"] as string)
                    };
                    FreeSql.Insert(req).ExecuteAffrows();
                    migrated++;
                }
                if (migrated > 0)
                {
                    _logger.Info($"旧数据迁移完成: 领退表 {migrated} 条");
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"旧数据迁移跳过: {ex.Message}");
            }
        }

        /// <summary>
        /// 兼容旧数据的日期解析：ISO 格式、中文月日（年份推断）、数字斜杠格式
        /// </summary>
        private static DateTime? ParseDate(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }
            if (DateTime.TryParseExact(text, ["yyyy-MM-dd HH:mm:ss.fff", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd", "yyyy/M/d", "M/d/yyyy"],
                System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime dt))
            {
                return dt;
            }
            System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(text, @"(\d{1,2})\s*月\s*(\d{1,2})\s*日");
            if (m.Success && int.TryParse(m.Groups[1].Value, out int month) && int.TryParse(m.Groups[2].Value, out int day))
            {
                try
                {
                    return new DateTime(DateTime.Now.Year, month, day);
                }
                catch
                {
                    return null;
                }
            }
            return DateTime.TryParse(text, out DateTime dt2) ? dt2 : null;
        }

        private static string NullIfEmpty(string s) => string.IsNullOrWhiteSpace(s) ? null : s.Trim();

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

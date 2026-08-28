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
            // 数据库路径可在设置中修改（保存在程序目录文件，重启生效）
            DbPath = AppSettingsService.ResolveDbPath();
            DataDir = Path.GetDirectoryName(DbPath);
            OleDir = Path.Combine(DataDir, "OleFiles");
            Directory.CreateDirectory(DataDir);
            Directory.CreateDirectory(OleDir);

            string connStr = $"Data Source={DbPath};Pooling=true;Min Pool Size=1;Max Pool Size=10";
            FreeSql = new FreeSqlBuilder()
                .UseConnectionString(DataType.Sqlite, connStr)
                .UseAutoSyncStructure(true) // 首次运行自动建表/同步结构
                .Build();

            // WAL 模式提升并发读写性能（最大并发10人次以内足够）
            FreeSql.Ado.ExecuteNonQuery("PRAGMA journal_mode=WAL;");
            MigrateLegacyPlansToRequisitions();
            _logger.Info($"数据库初始化完成: {DbPath}");
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

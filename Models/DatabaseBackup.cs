using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ORT一键报告.Models
{
    /// <summary>
    /// 备份类型：
    /// <see cref="Full"/> = 全量备份（整库快照，独立文件，可单独还原）；
    /// <see cref="Incremental"/> = 增量备份（只保存自上一个备份以来的数据变更，还原时先还原它的基准全量再依次回放增量）。
    /// </summary>
    public enum DatabaseBackupKind
    {
        Full,
        Incremental
    }

    /// <summary>
    /// 一份备份（全量快照文件或增量变更文件），供「快照与还原」列表与还原流程使用。
    /// </summary>
    public class DatabaseBackupEntry
    {
        /// <summary>备份文件完整路径</summary>
        public string FilePath { get; set; }

        /// <summary>备份文件名（还原链里用它作为标识）</summary>
        public string FileName { get; set; }

        /// <summary>全量 / 增量</summary>
        public DatabaseBackupKind Kind { get; set; }

        /// <summary>备份时间</summary>
        public DateTime CreatedAt { get; set; }

        /// <summary>文件大小（字节）</summary>
        public long SizeBytes { get; set; }

        /// <summary>增量备份的基准备份文件名（全量为空）</summary>
        public string BaseFileName { get; set; }

        /// <summary>增量包含的变更行数（新增+修改+删除）</summary>
        public int ChangedRows { get; set; }

        /// <summary>备份时的电脑名</summary>
        public string Machine { get; set; }

        /// <summary>备份时的操作者</summary>
        public string User { get; set; }

        /// <summary>备份时的程序版本</summary>
        public string AppVersion { get; set; }

        /// <summary>可否还原（文件在、基准链完整）</summary>
        public bool CanRestore { get; set; } = true;

        /// <summary>不可还原的原因（CanRestore 为 false 时给用户看）</summary>
        public string Problem { get; set; }

        /// <summary>文件大小的可读文本（KB/MB）</summary>
        public string SizeText => FormatSize(SizeBytes);

        public static string FormatSize(long bytes)
        {
            if (bytes < 1024)
            {
                return $"{bytes} B";
            }
            if (bytes < 1024 * 1024)
            {
                return $"{bytes / 1024.0:0.0} KB";
            }
            return $"{bytes / 1024.0 / 1024.0:0.0} MB";
        }
    }

    /// <summary>
    /// 增量备份里「一张表」的变更：新增/修改的整行 + 删除的键。
    /// </summary>
    public class DatabaseBackupTablePatch
    {
        /// <summary>表名</summary>
        public string Table { get; set; }

        /// <summary>用于定位行的键列（本程序各表都是自增主键 Id；没有整数主键时用 rowid）</summary>
        public string KeyColumn { get; set; } = "Id";

        /// <summary>键列是否为 SQLite 的 rowid 伪列</summary>
        public bool KeyIsRowId { get; set; }

        /// <summary>数据列（不含 rowid 伪列；含主键列，回放时整行覆盖）</summary>
        public List<string> Columns { get; set; } = [];

        /// <summary>被删除的行键</summary>
        public List<long> DeletedKeys { get; set; } = [];

        /// <summary>新增/修改的整行（键 + 与 <see cref="Columns"/> 一一对应的值）</summary>
        public List<DatabaseBackupRow> Rows { get; set; } = [];

        /// <summary>这张表的变更行数</summary>
        public int ChangedCount => DeletedKeys.Count + Rows.Count;
    }

    /// <summary>
    /// 增量备份里的一行数据：行键 + 整行取值（取值类型为 long/double/string/null）。
    /// </summary>
    public class DatabaseBackupRow
    {
        public long Key { get; set; }

        public object[] Values { get; set; } = [];
    }

    /// <summary>
    /// 增量备份文件的内容：只在基准备份之上记录变化的数据行，还原时按顺序回放即得到当时的数据库。
    /// </summary>
    public class DatabaseBackupPatch
    {
        /// <summary>固定为 Incremental（用于识别文件类型）</summary>
        public string Kind { get; set; } = "Incremental";

        /// <summary>备份时间</summary>
        public DateTime CreatedAt { get; set; }

        /// <summary>基准备份文件名（还原时先还原它，再依次回放之后的增量）</summary>
        public string BaseFileName { get; set; }

        /// <summary>数据库文件名</summary>
        public string DbFileName { get; set; }

        /// <summary>备份时的电脑名</summary>
        public string Machine { get; set; }

        /// <summary>备份时的操作者</summary>
        public string User { get; set; }

        /// <summary>备份时的程序版本</summary>
        public string AppVersion { get; set; }

        /// <summary>各表的变更</summary>
        public List<DatabaseBackupTablePatch> Tables { get; set; } = [];

        /// <summary>本次增量共包含多少行变更</summary>
        public int ChangedRows => Tables?.Sum(t => t.ChangedCount) ?? 0;
    }

    /// <summary>
    /// 某张表在某个时刻的完整快照（内存里用于比对差异）：
    /// 键 → 整行取值，配上列名，即可与另一时刻的快照比出新增/修改/删除。
    /// </summary>
    public class DatabaseBackupTableSnapshot
    {
        public string Table { get; set; }

        /// <summary>键列名（Id 或 rowid）</summary>
        public string KeyColumn { get; set; } = "Id";

        /// <summary>键列是否为 rowid 伪列</summary>
        public bool KeyIsRowId { get; set; }

        /// <summary>数据列（不含 rowid 伪列）</summary>
        public List<string> Columns { get; set; } = [];

        /// <summary>键 → 整行取值</summary>
        public Dictionary<long, object[]> Rows { get; set; } = [];

        /// <summary>取某行的归一化签名（用于判断行内容是否变化）</summary>
        public string SignatureOf(long key)
            => Rows.TryGetValue(key, out object[] values) ? DatabaseBackupDiffer.RowSignature(values) : null;
    }

    /// <summary>
    /// 增量比对与备份文件名解析（纯函数，便于单元测试）：
    /// 把「基准快照」与「当前快照」比成一份可回放的变更补丁。
    /// </summary>
    public static class DatabaseBackupDiffer
    {
        /// <summary>全量备份文件名前缀</summary>
        public const string FullPrefix = "ort_plans_full_";

        /// <summary>增量备份文件名前缀</summary>
        public const string IncrementalPrefix = "ort_plans_inc_";

        /// <summary>文件名里的时间戳格式（与文件名前缀一起解析）</summary>
        public const string TimeFormat = "yyyyMMdd_HHmmss";

        /// <summary>BLOB 值在补丁里的前缀（本项目各表没有二进制列，留作兜底）</summary>
        private const char BlobMarker = '\u0001';

        /// <summary>
        /// 比较两张表的快照，得到这张表的变更（只在 <paramref name="after"/> 里出现的键视为新增/修改，只在 before 里出现的视为删除）。
        /// <paramref name="before"/> 为 null（新表）时整表内容都算新增。
        /// </summary>
        public static DatabaseBackupTablePatch DiffTable(DatabaseBackupTableSnapshot before, DatabaseBackupTableSnapshot after)
        {
            if (after == null)
            {
                return null;
            }
            DatabaseBackupTablePatch patch = new()
            {
                Table = after.Table,
                KeyColumn = after.KeyColumn,
                KeyIsRowId = after.KeyIsRowId,
                Columns = [.. after.Columns]
            };
            foreach (KeyValuePair<long, object[]> row in after.Rows)
            {
                string oldSignature = before?.SignatureOf(row.Key);
                if (oldSignature == null || !string.Equals(oldSignature, RowSignature(row.Value), StringComparison.Ordinal))
                {
                    patch.Rows.Add(new DatabaseBackupRow { Key = row.Key, Values = row.Value });
                }
            }
            if (before != null)
            {
                foreach (long key in before.Rows.Keys)
                {
                    if (!after.Rows.ContainsKey(key))
                    {
                        patch.DeletedKeys.Add(key);
                    }
                }
            }
            return patch;
        }

        /// <summary>
        /// 整行的归一化签名：null、字符串、整数、浮点、BLOB 分开编码，避免不同类型被比成同一个值
        /// </summary>
        public static string RowSignature(object[] values)
        {
            if (values == null)
            {
                return "";
            }
            return string.Join("\u0002", values.Select(ValueSignature));
        }

        /// <summary>单个取值的归一化文本（整数与数字文本分开编码，避免 1 与 "1" 被当成同一个值）</summary>
        public static string ValueSignature(object value)
        {
            switch (value)
            {
                case null:
                case DBNull:
                    return "\u0000";
                case byte[] blob:
                    return "B:" + Convert.ToBase64String(blob);
                case bool b:
                    return "L:" + (b ? "1" : "0");
                case long l:
                    return "I:" + l.ToString(System.Globalization.CultureInfo.InvariantCulture);
                case int i:
                    return "I:" + i.ToString(System.Globalization.CultureInfo.InvariantCulture);
                case short s:
                    return "I:" + s.ToString(System.Globalization.CultureInfo.InvariantCulture);
                case byte by:
                    return "I:" + by.ToString(System.Globalization.CultureInfo.InvariantCulture);
                case sbyte sb:
                    return "I:" + sb.ToString(System.Globalization.CultureInfo.InvariantCulture);
                case uint ui:
                    return "I:" + ui.ToString(System.Globalization.CultureInfo.InvariantCulture);
                case ulong ul:
                    return "I:" + ul.ToString(System.Globalization.CultureInfo.InvariantCulture);
                case ushort us:
                    return "I:" + us.ToString(System.Globalization.CultureInfo.InvariantCulture);
                case double d:
                    return "R:" + d.ToString("R", System.Globalization.CultureInfo.InvariantCulture);
                case float f:
                    return "R:" + ((double)f).ToString("R", System.Globalization.CultureInfo.InvariantCulture);
                case decimal m:
                    return "N:" + m.ToString(System.Globalization.CultureInfo.InvariantCulture);
                default:
                    return "S:" + Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture);
            }
        }

        /// <summary>取值的字面量（数字/BLOB 原样、文本加引号并按 SQL 规则转义）</summary>
        public static string ValueLiteral(object value)
        {
            switch (value)
            {
                case null:
                case DBNull:
                    return "NULL";
                case byte[] blob:
                    return "X'" + ToHex(blob) + "'";
                case long l:
                    return l.ToString(System.Globalization.CultureInfo.InvariantCulture);
                case int i:
                    return i.ToString(System.Globalization.CultureInfo.InvariantCulture);
                case short s:
                    return s.ToString(System.Globalization.CultureInfo.InvariantCulture);
                case byte by:
                    return by.ToString(System.Globalization.CultureInfo.InvariantCulture);
                case sbyte sb:
                    return sb.ToString(System.Globalization.CultureInfo.InvariantCulture);
                case uint ui:
                    return ui.ToString(System.Globalization.CultureInfo.InvariantCulture);
                case ulong ul:
                    return ul.ToString(System.Globalization.CultureInfo.InvariantCulture);
                case ushort us:
                    return us.ToString(System.Globalization.CultureInfo.InvariantCulture);
                case double d:
                    return d.ToString("R", System.Globalization.CultureInfo.InvariantCulture);
                case float f:
                    return ((double)f).ToString("R", System.Globalization.CultureInfo.InvariantCulture);
                case decimal m:
                    return m.ToString(System.Globalization.CultureInfo.InvariantCulture);
                case bool b:
                    return b ? "1" : "0";
                case string text when text.Length > 0 && text[0] == BlobMarker:
                    // 快照里把 BLOB 编成了标记 + Base64，回放时还原成二进制
                    return "X'" + ToHex(Convert.FromBase64String(text.Substring(1))) + "'";
                default:
                    return "'" + Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture).Replace("'", "''") + "'";
            }
        }

        /// <summary>二进制转十六进制文本（.NET Framework 上没有 Convert.ToHexString）</summary>
        private static string ToHex(byte[] bytes) => BitConverter.ToString(bytes).Replace("-", "");

        /// <summary>把取值规整成可 JSON 往返的类型（BLOB 转成标记 + Base64 文本）</summary>
        public static object ValueForStorage(object value)
        {
            if (value == null || value is DBNull)
            {
                return null;
            }
            return value is byte[] blob ? BlobMarker + Convert.ToBase64String(blob) : value;
        }

        /* ###############################  文件名  ################################ */

        /// <summary>
        /// 从备份文件名解析类型与时间（认不出返回 false）。
        /// 形如 ort_plans_full_20261009_153000.db / ort_plans_inc_20261009_153000.json。
        /// </summary>
        public static bool TryParseFileName(string fileName, out DatabaseBackupKind kind, out DateTime createdAt)
        {
            kind = DatabaseBackupKind.Full;
            createdAt = default;
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return false;
            }
            fileName = Path.GetFileName(fileName);
            string rest;
            if (fileName.StartsWith(FullPrefix, StringComparison.OrdinalIgnoreCase))
            {
                kind = DatabaseBackupKind.Full;
                rest = fileName.Substring(FullPrefix.Length);
            }
            else if (fileName.StartsWith(IncrementalPrefix, StringComparison.OrdinalIgnoreCase))
            {
                kind = DatabaseBackupKind.Incremental;
                rest = fileName.Substring(IncrementalPrefix.Length);
            }
            else
            {
                return false;
            }
            int dot = rest.IndexOf('.');
            if (dot >= 0)
            {
                rest = rest.Substring(0, dot);
            }
            return DateTime.TryParseExact(rest, TimeFormat, System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out createdAt);
        }

        /// <summary>全量备份文件名（时间戳秒级，同一秒内不会重复生成）</summary>
        public static string FullFileName(DateTime time) => $"{FullPrefix}{time.ToString(TimeFormat)}.db";

        /// <summary>增量备份文件名</summary>
        public static string IncrementalFileName(DateTime time) => $"{IncrementalPrefix}{time.ToString(TimeFormat)}.json";

        /* ###############################  还原链  ################################ */

        /// <summary>
        /// 组装还原链：从最近的全量开始，沿增量备份的基准链接一路走到 <paramref name="target"/>。
        /// 链上缺文件（被删或改名）时返回 null，并给出 <paramref name="problem"/>。
        /// </summary>
        public static List<DatabaseBackupEntry> BuildRestoreChain(
            IReadOnlyDictionary<string, DatabaseBackupEntry> byFileName, DatabaseBackupEntry target, out string problem)
        {
            problem = null;
            if (target == null)
            {
                problem = "未选择备份";
                return null;
            }
            List<DatabaseBackupEntry> chain = [target];
            HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase) { target.FileName };
            DatabaseBackupEntry current = target;
            while (current.Kind == DatabaseBackupKind.Incremental)
            {
                if (string.IsNullOrWhiteSpace(current.BaseFileName))
                {
                    problem = $"{current.FileName} 没有记录基准备份，无法还原";
                    return null;
                }
                if (!byFileName.TryGetValue(current.BaseFileName, out DatabaseBackupEntry baseEntry))
                {
                    problem = $"{current.FileName} 的基准备份 {current.BaseFileName} 不在备份文件夹里，无法还原";
                    return null;
                }
                if (!seen.Add(baseEntry.FileName))
                {
                    problem = $"{current.FileName} 的基准链出现循环，无法还原";
                    return null;
                }
                chain.Add(baseEntry);
                current = baseEntry;
            }
            chain.Reverse();
            return chain;
        }
    }
}

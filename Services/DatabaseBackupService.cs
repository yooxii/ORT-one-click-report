using FreeSql;
using NLog;
using Newtonsoft.Json;
using ORT一键报告.Models;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 备份操作结果（给界面直接显示的中文消息）。
    /// </summary>
    public class DatabaseBackupResult
    {
        /// <summary>操作是否成功（失败时 Message 是给用户看的原因）</summary>
        public bool Success { get; set; }

        /// <summary>是否真的生成了备份文件（条件不满足时为 false：已有本周全量/今日增量）</summary>
        public bool Created { get; set; }

        /// <summary>界面提示消息</summary>
        public string Message { get; set; }

        /// <summary>生成的（或相关的那份）备份</summary>
        public DatabaseBackupEntry Entry { get; set; }

        /// <summary>表结构变了、增量做不了，建议改做全量</summary>
        public bool RequiresFullBackup { get; set; }

        public static DatabaseBackupResult Ok(string message, bool created, DatabaseBackupEntry entry = null)
            => new() { Success = true, Created = created, Message = message, Entry = entry };

        public static DatabaseBackupResult Fail(string message)
            => new() { Success = false, Message = message };
    }

    /// <summary>
    /// 数据库备份与还原：
    /// <list type="bullet">
    /// <item>全量备份 = 用 SQLite 的 VACUUM INTO 生成一份一致的整库快照文件（.db），可独立还原；</item>
    /// <item>增量备份 = 与上一个备份逐表逐行比对，只把新增/修改/删除的数据行写成补丁文件（.json），
    /// 还原时先还原它的基准全量、再按时间顺序回放增量；</item>
    /// <item>默认备份目录 = 数据库同级目录下的 Backups（可在设置里改），多台电脑指向同一数据库时靠
    /// 目录里的锁文件与「本周已有全量 / 今日已有增量」判断避免重复备份。</item>
    /// </list>
    /// 只备份数据库文件本身；附件目录（OleFiles）与配图目录（PlanImages）不在备份范围内。
    /// </summary>
    public class DatabaseBackupService
    {
        /// <summary>是否启用定时备份（app_settings 键，默认启用）</summary>
        public const string SettingAutoKey = "backup.auto";

        /// <summary>备份文件夹（app_settings 键；为空表示数据库同级的 Backups）</summary>
        public const string SettingFolderKey = "backup.folder";

        /// <summary>全量备份间隔（天）：距上次全量超过这个天数就再做一次全量</summary>
        public const int FullBackupIntervalDays = 7;

        /// <summary>默认备份文件夹名（数据库同级）</summary>
        public const string DefaultFolderName = "Backups";

        private const string LockFileName = ".ort_backup.lock";
        private const string RestoreTempName = ".ort_restore_tmp.db";

        /// <summary>锁文件超过这个时长视为上次备份异常中断留下的残留</summary>
        private static readonly TimeSpan LockStaleAfter = TimeSpan.FromMinutes(30);

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;
        private readonly AppSettingsService _settings;

        public DatabaseBackupService(DatabaseService db, AppSettingsService settings)
        {
            _db = db;
            _settings = settings;
        }

        /// <summary>当前生效的备份文件夹（设置里填了就用它，否则数据库同级的 Backups）</summary>
        public string ResolveBackupRoot()
        {
            string configured = FolderUtil.Normalize(_settings.GetText(SettingFolderKey));
            return string.IsNullOrEmpty(configured)
                ? Path.Combine(_db.DataDir, DefaultFolderName)
                : configured;
        }

        /// <summary>默认备份文件夹（数据库同级的 Backups），设置界面「恢复默认」用</summary>
        public string DefaultBackupRoot => Path.Combine(_db.DataDir, DefaultFolderName);

        /// <summary>是否启用定时备份（默认启用）</summary>
        public bool IsAutoEnabled => _settings.GetBool(SettingAutoKey, true);

        /* ###############################  定时策略  ################################ */

        /// <summary>
        /// 按定时策略执行一次：距上次全量满 7 天（或还没有全量）→ 做全量；否则当天还没有增量 → 做增量；
        /// 两者都不需要时什么都不做。多台电脑同时到点靠备份目录里的锁文件互斥，只有一台真正执行。
        /// </summary>
        public DatabaseBackupResult RunScheduled()
        {
            if (!IsAutoEnabled)
            {
                return DatabaseBackupResult.Ok("定时备份未启用", false);
            }
            string root = ResolveBackupRoot();
            if (!FolderUtil.TryPrepare(root, out string folderError))
            {
                return DatabaseBackupResult.Fail($"备份文件夹不可用：{folderError}");
            }
            List<DatabaseBackupEntry> entries = ListBackups();
            DatabaseBackupEntry latestFull = entries.FirstOrDefault(e => e.Kind == DatabaseBackupKind.Full);
            DateTime now = DateTime.Now;
            if (latestFull == null || (now - latestFull.CreatedAt).TotalDays >= FullBackupIntervalDays)
            {
                string reason = latestFull == null
                    ? "定时全量备份（首次）"
                    : $"定时全量备份（距上次全量 {(now - latestFull.CreatedAt).TotalDays:0.0} 天）";
                return CreateFullBackup(reason);
            }
            bool incrementalToday = entries.Any(e => e.Kind == DatabaseBackupKind.Incremental
                && e.CreatedAt.Date == now.Date);
            if (incrementalToday)
            {
                return DatabaseBackupResult.Ok("今天的增量备份已经做过", false);
            }
            return CreateIncrementalBackup("定时增量备份（每日一次）");
        }

        /* ###############################  生成备份  ################################ */

        /// <summary>
        /// 生成全量备份（整库快照，VACUUM INTO 保证文件一致、不含 WAL 残留）
        /// </summary>
        public DatabaseBackupResult CreateFullBackup(string reason)
        {
            string root = ResolveBackupRoot();
            if (!FolderUtil.TryPrepare(root, out string folderError))
            {
                return DatabaseBackupResult.Fail($"备份文件夹不可用：{folderError}");
            }
            using FileStream lockFile = TryAcquireLock(root);
            if (lockFile == null)
            {
                return DatabaseBackupResult.Ok("另一个备份任务正在执行，本次跳过", false);
            }
            return CreateFullBackupCore(root, reason);
        }

        /// <summary>生成全量备份的实际动作（调用方必须已经拿到备份锁）</summary>
        private DatabaseBackupResult CreateFullBackupCore(string root, string reason)
        {
            string fileName = DatabaseBackupDiffer.FullFileName(DateTime.Now);
            string target = Path.Combine(root, fileName);
            string temp = target + ".tmp";
            try
            {
                TryDelete(temp);
                // VACUUM INTO：由 SQLite 自己生成一份一致的库文件（WAL 里的内容也会被合并进去）
                _db.FreeSql.Ado.ExecuteNonQuery($"VACUUM INTO '{EscapeSql(target)}'");
                TryDelete(target);
                File.Move(temp, target);
                DatabaseBackupEntry entry = Describe(target, DatabaseBackupKind.Full);
                entry.Machine = Environment.MachineName;
                entry.User = Environment.UserName;
                entry.AppVersion = AppVersion;
                _logger.Info($"全量备份完成（{reason}）：{target}（{entry.SizeText}）");
                return DatabaseBackupResult.Ok($"全量备份完成：{fileName}（{entry.SizeText}）", true, entry);
            }
            catch (Exception ex)
            {
                TryDelete(temp);
                _logger.Error(ex, $"全量备份失败（{reason}）");
                return DatabaseBackupResult.Fail($"全量备份失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 生成增量备份：与最近的备份逐表逐行比对，只把变化的数据行写成补丁文件。
        /// 没有任何备份时先做全量；表结构（表/列）与基准不一致时改为做全量。
        /// </summary>
        public DatabaseBackupResult CreateIncrementalBackup(string reason)
        {
            string root = ResolveBackupRoot();
            if (!FolderUtil.TryPrepare(root, out string folderError))
            {
                return DatabaseBackupResult.Fail($"备份文件夹不可用：{folderError}");
            }
            List<DatabaseBackupEntry> entries = ListBackups();
            DatabaseBackupEntry baseEntry = entries.FirstOrDefault();
            if (baseEntry == null)
            {
                return CreateFullBackup($"{reason}（还没有基准备份，先做全量）");
            }
            using FileStream lockFile = TryAcquireLock(root);
            if (lockFile == null)
            {
                return DatabaseBackupResult.Ok("另一个备份任务正在执行，本次跳过", false);
            }
            try
            {
                DateTime now = DateTime.Now;
                DatabaseBackupPatch patch = new()
                {
                    CreatedAt = now,
                    BaseFileName = baseEntry.FileName,
                    DbFileName = Path.GetFileName(_db.DbPath),
                    Machine = Environment.MachineName,
                    User = Environment.UserName,
                    AppVersion = AppVersion
                };
                IFreeSql baseFsql = null;
                try
                {
                    baseFsql = OpenReadOnly(baseEntry.FilePath);
                    Dictionary<string, DatabaseBackupTableSnapshot> before = ReadSnapshot(baseFsql);
                    Dictionary<string, DatabaseBackupTableSnapshot> after = ReadSnapshot(_db.FreeSql);
                    if (!SameSchema(before, after))
                    {
                        _logger.Info("数据库表结构与基准备份不一致（新增了表/列），本次改为全量备份");
                        return CreateFullBackupCore(root, $"{reason}（表结构已变化）");
                    }
                    foreach (KeyValuePair<string, DatabaseBackupTableSnapshot> table in after)
                    {
                        DatabaseBackupTablePatch tablePatch = DatabaseBackupDiffer.DiffTable(
                            before.TryGetValue(table.Key, out DatabaseBackupTableSnapshot old) ? old : null,
                            table.Value);
                        if (tablePatch != null && tablePatch.ChangedCount > 0)
                        {
                            patch.Tables.Add(tablePatch);
                        }
                    }
                }
                catch (InvalidOperationException ex)
                {
                    // 出现不认识的表结构（非整数自增主键等）：增量做不了，退化为全量
                    _logger.Warn($"增量比对不可用（{ex.Message}），本次改为全量备份");
                    return CreateFullBackupCore(root, $"{reason}（无法逐行比对）");
                }
                finally
                {
                    baseFsql?.Dispose();
                    ClearSqlitePools();
                }

                string fileName = DatabaseBackupDiffer.IncrementalFileName(now);
                string target = Path.Combine(root, fileName);
                string temp = target + ".tmp";
                TryDelete(temp);
                File.WriteAllText(temp, JsonConvert.SerializeObject(patch), new System.Text.UTF8Encoding(false));
                TryDelete(target);
                File.Move(temp, target);
                DatabaseBackupEntry entry = Describe(target, DatabaseBackupKind.Incremental);
                entry.BaseFileName = patch.BaseFileName;
                entry.ChangedRows = patch.ChangedRows;
                entry.Machine = patch.Machine;
                entry.User = patch.User;
                entry.AppVersion = patch.AppVersion;
                _logger.Info($"增量备份完成（{reason}）：{target}（基准 {patch.BaseFileName}，变更 {patch.ChangedRows} 行，{entry.SizeText}）");
                return DatabaseBackupResult.Ok($"增量备份完成：{fileName}（变更 {patch.ChangedRows} 行，{entry.SizeText}）", true, entry);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"增量备份失败（{reason}）");
                return DatabaseBackupResult.Fail($"增量备份失败：{ex.Message}");
            }
        }

        /* ###############################  快照列表  ################################ */

        /// <summary>
        /// 列出备份文件夹里的全部备份（新的在前）；增量会读出基准文件名与变更行数，
        /// 并标出哪些因为基准缺失而不能还原。
        /// </summary>
        public List<DatabaseBackupEntry> ListBackups()
        {
            List<DatabaseBackupEntry> result = [];
            string root = ResolveBackupRoot();
            if (!Directory.Exists(root))
            {
                return result;
            }
            Dictionary<string, DatabaseBackupEntry> byName = new(StringComparer.OrdinalIgnoreCase);
            foreach (string file in Directory.EnumerateFiles(root))
            {
                string name = Path.GetFileName(file);
                if (!DatabaseBackupDiffer.TryParseFileName(name, out DatabaseBackupKind kind, out DateTime createdAt))
                {
                    continue;
                }
                DatabaseBackupEntry entry = Describe(file, kind);
                entry.CreatedAt = createdAt;
                if (kind == DatabaseBackupKind.Incremental)
                {
                    DatabaseBackupPatch patch = TryReadPatch(file);
                    if (patch == null)
                    {
                        entry.CanRestore = false;
                        entry.Problem = "文件读取失败或不是有效的增量备份";
                    }
                    else
                    {
                        entry.BaseFileName = patch.BaseFileName;
                        entry.ChangedRows = patch.ChangedRows;
                        entry.Machine = patch.Machine;
                        entry.User = patch.User;
                        entry.AppVersion = patch.AppVersion;
                        entry.CreatedAt = patch.CreatedAt == default ? createdAt : patch.CreatedAt;
                    }
                }
                result.Add(entry);
                byName[entry.FileName] = entry;
            }
            result.Sort((a, b) => b.CreatedAt.CompareTo(a.CreatedAt));
            // 标注基准链缺失的增量：不能单独还原
            foreach (DatabaseBackupEntry entry in result.Where(e => e.Kind == DatabaseBackupKind.Incremental && e.CanRestore))
            {
                if (DatabaseBackupDiffer.BuildRestoreChain(byName, entry, out string problem) == null)
                {
                    entry.CanRestore = false;
                    entry.Problem = problem;
                }
            }
            return result;
        }

        /// <summary>把备份文件夹里的备份按文件名索引（还原链组装用）</summary>
        public Dictionary<string, DatabaseBackupEntry> IndexBackups(IEnumerable<DatabaseBackupEntry> entries)
            => (entries ?? []).Where(e => e != null).GroupBy(e => e.FileName, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        /// <summary>删除一份备份文件</summary>
        public DatabaseBackupResult DeleteBackup(DatabaseBackupEntry entry)
        {
            try
            {
                if (entry == null || !File.Exists(entry.FilePath))
                {
                    return DatabaseBackupResult.Fail("备份文件不存在");
                }
                File.Delete(entry.FilePath);
                _logger.Info($"已删除备份：{entry.FilePath}");
                return DatabaseBackupResult.Ok($"已删除 {entry.FileName}", false);
            }
            catch (Exception ex)
            {
                return DatabaseBackupResult.Fail($"删除失败：{ex.Message}");
            }
        }

        /* ###############################  还原  ################################ */

        /// <summary>
        /// 还原到指定备份：先给当前库做一次全量备份（还原前的保底），
        /// 再把「基准全量 + 之后的增量」回放成一份数据库文件，最后替换正在使用的库文件。
        /// 替换成功后本进程的数据库连接已不可用，调用方必须重启程序。
        /// </summary>
        public DatabaseBackupResult Restore(DatabaseBackupEntry target, IProgress<string> progress)
        {
            List<DatabaseBackupEntry> entries = ListBackups();
            Dictionary<string, DatabaseBackupEntry> byName = IndexBackups(entries);
            List<DatabaseBackupEntry> chain = DatabaseBackupDiffer.BuildRestoreChain(
                byName, byName.TryGetValue(target?.FileName ?? "", out DatabaseBackupEntry fresh) ? fresh : target, out string problem);
            if (chain == null)
            {
                return DatabaseBackupResult.Fail($"无法还原：{problem}");
            }
            DatabaseBackupEntry full = chain[0];
            if (!File.Exists(full.FilePath))
            {
                return DatabaseBackupResult.Fail($"无法还原：基准全量备份 {full.FileName} 已不在备份文件夹里");
            }

            string root = ResolveBackupRoot();
            if (!FolderUtil.TryPrepare(root, out string folderError))
            {
                return DatabaseBackupResult.Fail($"无法还原：备份文件夹不可用（{folderError}）");
            }
            // 整个还原过程独占备份目录里的锁：期间不允许其它备份/还原插进来
            using FileStream lockFile = TryAcquireLock(root);
            if (lockFile == null)
            {
                return DatabaseBackupResult.Fail("无法还原：另一个备份或还原任务正在执行，请稍后再试");
            }

            // 1) 还原前先给当前库做一次全量备份：还原是破坏性操作，保底能退回来
            progress?.Report("正在备份当前数据库…");
            DatabaseBackupResult safety = CreateFullBackupCore(root, "还原前自动备份");
            if (!safety.Success || !safety.Created)
            {
                return DatabaseBackupResult.Fail($"还原已取消：还原前的保底备份没有成功（{safety.Message}）");
            }

            string temp = Path.Combine(root, RestoreTempName);
            try
            {
                // 2) 基准全量复制成临时文件，再依次回放增量
                progress?.Report("正在展开快照…");
                TryDelete(temp);
                ClearSqlitePools();
                File.Copy(full.FilePath, temp, true);
                IFreeSql work = null;
                try
                {
                    work = OpenReadOnly(temp, readOnly: false);
                    for (int i = 1; i < chain.Count; i++)
                    {
                        DatabaseBackupEntry step = chain[i];
                        progress?.Report($"正在回放增量备份 {step.FileName}…");
                        DatabaseBackupPatch patch = TryReadPatch(step.FilePath);
                        if (patch == null)
                        {
                            return DatabaseBackupResult.Fail($"无法还原：增量备份 {step.FileName} 读取失败");
                        }
                        ApplyPatch(work, patch, step.FileName);
                    }
                    progress?.Report("正在校验还原结果…");
                    string check = Convert.ToString(work.Ado.ExecuteScalar("PRAGMA integrity_check;"), CultureInfo.InvariantCulture);
                    if (!string.Equals(check, "ok", StringComparison.OrdinalIgnoreCase))
                    {
                        return DatabaseBackupResult.Fail($"还原失败：还原后的数据库未通过完整性检查（{check}）");
                    }
                }
                finally
                {
                    work?.Dispose();
                    ClearSqlitePools();
                    TryDelete(temp + "-wal");
                    TryDelete(temp + "-shm");
                    TryDelete(temp + "-journal");
                }

                // 3) 替换正在使用的数据库文件（连接已被释放；失败会把原库改回名字，不会两头空）
                progress?.Report("正在替换数据库文件…");
                _db.ReplaceDatabaseFile(temp);
                _logger.Warn($"数据库已还原到 {target.FileName}（基准 {full.FileName}，增量 {chain.Count - 1} 份），需要重启程序");
                return DatabaseBackupResult.Ok(
                    $"已还原到 {target.CreatedAt:yyyy/M/d HH:mm:ss} 的{(target.Kind == DatabaseBackupKind.Full ? "全量" : "增量")}备份（还原前已自动备份当前数据库）",
                    true, target);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"还原到 {target?.FileName} 失败");
                TryDelete(temp);
                return DatabaseBackupResult.Fail($"还原失败：{ex.Message}");
            }
            finally
            {
                TryDelete(temp);
            }
        }

        /// <summary>
        /// 把一份增量补丁回放到目标库：逐表整行覆盖（INSERT OR REPLACE）+ 删除。
        /// 中途失败直接抛出——调用方用的是临时文件，丢弃即可，不会影响正式库。
        /// </summary>
        private static void ApplyPatch(IFreeSql work, DatabaseBackupPatch patch, string fileName)
        {
            HashSet<string> tables = ReadTableNames(work);
            foreach (DatabaseBackupTablePatch table in patch.Tables ?? [])
            {
                if (!tables.Contains(table.Table))
                {
                    throw new InvalidOperationException($"{fileName} 里的表 {table.Table} 在基准备份中不存在");
                }
                foreach (long key in table.DeletedKeys ?? [])
                {
                    work.Ado.ExecuteNonQuery($"DELETE FROM \"{table.Table}\" WHERE \"{table.KeyColumn}\" = {key};");
                }
                if (table.Rows == null || table.Rows.Count == 0)
                {
                    continue;
                }
                string columns = string.Join(",", table.Columns.Select(c => $"\"{c}\""));
                foreach (DatabaseBackupRow row in table.Rows)
                {
                    object[] values = row.Values ?? [];
                    if (values.Length != table.Columns.Count)
                    {
                        throw new InvalidOperationException($"{fileName} 里的表 {table.Table} 行数据列数与列定义不一致");
                    }
                    string literals = string.Join(",", values.Select(DatabaseBackupDiffer.ValueLiteral));
                    work.Ado.ExecuteNonQuery($"INSERT OR REPLACE INTO \"{table.Table}\" ({columns}) VALUES ({literals});");
                }
            }
        }

        /* ###############################  快照读取  ################################ */

        private Dictionary<string, DatabaseBackupTableSnapshot> ReadSnapshot(IFreeSql fsql)
        {
            Dictionary<string, DatabaseBackupTableSnapshot> snapshot = new(StringComparer.Ordinal);
            foreach (string table in ReadTableNames(fsql))
            {
                // 只处理「单个整数主键（自增 Id）」的表：本程序所有业务表都是这个形状，
                // 其它形状的表（理论上不会出现）直接抛出让上层改做全量备份，避免写坏数据
                DatabaseBackupTableSnapshot tableSnapshot = ReadTableSnapshot(fsql, table);
                snapshot[table] = tableSnapshot;
            }
            return snapshot;
        }

        private static HashSet<string> ReadTableNames(IFreeSql fsql)
        {
            DataTable table = fsql.Ado.ExecuteDataTable(
                "SELECT name FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%' ORDER BY name;");
            HashSet<string> names = new(StringComparer.Ordinal);
            foreach (DataRow row in table.Rows)
            {
                names.Add(Convert.ToString(row[0], CultureInfo.InvariantCulture));
            }
            return names;
        }

        private static DatabaseBackupTableSnapshot ReadTableSnapshot(IFreeSql fsql, string table)
        {
            string keyColumn = null;
            DataTable info = fsql.Ado.ExecuteDataTable($"SELECT name, type, pk FROM pragma_table_info('{EscapeSql(table)}');");
            foreach (DataRow row in info.Rows)
            {
                if (Convert.ToInt32(row["pk"], CultureInfo.InvariantCulture) > 0)
                {
                    string name = Convert.ToString(row["name"], CultureInfo.InvariantCulture);
                    string type = Convert.ToString(row["type"], CultureInfo.InvariantCulture) ?? "";
                    if (name.Any(char.IsWhiteSpace) || !type.Contains("INT", StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidOperationException($"表 {table} 的主键不是整数自增列，无法做增量备份");
                    }
                    keyColumn = name;
                }
            }
            if (keyColumn == null)
            {
                throw new InvalidOperationException($"表 {table} 没有主键，无法做增量备份");
            }

            DataTable data = fsql.Ado.ExecuteDataTable($"SELECT * FROM \"{table}\";");
            DatabaseBackupTableSnapshot snapshot = new()
            {
                Table = table,
                KeyColumn = keyColumn,
                KeyIsRowId = false,
                Columns = data.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList()
            };
            int keyIndex = snapshot.Columns.IndexOf(keyColumn);
            foreach (DataRow row in data.Rows)
            {
                long key = Convert.ToInt64(row[keyIndex], CultureInfo.InvariantCulture);
                object[] values = new object[snapshot.Columns.Count];
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = DatabaseBackupDiffer.ValueForStorage(row[i]);
                }
                snapshot.Rows[key] = values;
            }
            return snapshot;
        }

        /// <summary>表集合与列集合是否一致（不一致就没法用增量补丁回放，改做全量）</summary>
        private static bool SameSchema(Dictionary<string, DatabaseBackupTableSnapshot> before,
            Dictionary<string, DatabaseBackupTableSnapshot> after)
        {
            if (before.Count != after.Count)
            {
                return false;
            }
            foreach (KeyValuePair<string, DatabaseBackupTableSnapshot> table in after)
            {
                if (!before.TryGetValue(table.Key, out DatabaseBackupTableSnapshot old))
                {
                    return false;
                }
                if (!old.Columns.SequenceEqual(table.Value.Columns, StringComparer.Ordinal))
                {
                    return false;
                }
            }
            return true;
        }

        /* ###############################  基础设施  ################################ */

        /// <summary>按文件信息构造备份条目（增量的基准/变更行数由调用方或 ListBackups 补）</summary>
        private static DatabaseBackupEntry Describe(string filePath, DatabaseBackupKind kind)
        {
            FileInfo info = new(filePath);
            return new DatabaseBackupEntry
            {
                FilePath = filePath,
                FileName = Path.GetFileName(filePath),
                Kind = kind,
                CreatedAt = DatabaseBackupDiffer.TryParseFileName(info.Name, out _, out DateTime created) ? created : info.LastWriteTime,
                SizeBytes = info.Exists ? info.Length : 0
            };
        }

        private static DatabaseBackupPatch TryReadPatch(string filePath)
        {
            try
            {
                DatabaseBackupPatch patch = JsonConvert.DeserializeObject<DatabaseBackupPatch>(File.ReadAllText(filePath));
                return string.Equals(patch?.Kind, "Incremental", StringComparison.OrdinalIgnoreCase) ? patch : null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 打开一份备份库文件（Pooling=false：用完立即释放文件句柄，便于删除/替换该文件）
        /// </summary>
        private static IFreeSql OpenReadOnly(string filePath, bool readOnly = true)
        {
            string mode = readOnly ? "Read Only=True;" : "";
            return new FreeSqlBuilder()
                .UseConnectionString(DataType.Sqlite, $"Data Source={filePath};Pooling=false;{mode}")
                .Build();
        }

        /// <summary>
        /// 释放 ADO.NET 连接池里可能还握着的 SQLite 文件句柄：
        /// 替换数据库文件前必须确保没有任何连接占着它。用反射调用，避免程序集对本项目产生硬依赖。
        /// </summary>
        public static void ClearSqlitePools()
        {
            try
            {
                Type type = Type.GetType("System.Data.SQLite.SQLiteConnection, System.Data.SQLite");
                type?.GetMethod("ClearAllPools", BindingFlags.Public | BindingFlags.Static)?.Invoke(null, null);
            }
            catch (Exception ex)
            {
                LogManager.GetCurrentClassLogger().Warn($"清理 SQLite 连接池失败: {ex.Message}");
            }
        }

        /// <summary>备份目录里的互斥锁（多台电脑同时到点时只让一台执行）；拿不到返回 null</summary>
        private static FileStream TryAcquireLock(string root)
        {
            string path = Path.Combine(root, LockFileName);
            for (int attempt = 0; attempt < 2; attempt++)
            {
                try
                {
                    return new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
                }
                catch (IOException)
                {
                    try
                    {
                        if (attempt == 0 && File.Exists(path)
                            && DateTime.Now - File.GetLastWriteTime(path) > LockStaleAfter)
                        {
                            File.Delete(path);   // 上次备份异常中断留下的残留锁
                            continue;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.GetCurrentClassLogger().Warn($"清理残留备份锁失败: {ex.Message}");
                    }
                    return null;
                }
                catch (Exception ex)
                {
                    LogManager.GetCurrentClassLogger().Warn($"获取备份锁失败: {ex.Message}");
                    return null;
                }
            }
            return null;
        }

        private static void TryDelete(string path)
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
                LogManager.GetCurrentClassLogger().Warn($"删除文件失败（{path}）: {ex.Message}");
            }
        }

        private static string EscapeSql(string text) => text?.Replace("'", "''");

        /// <summary>当前程序版本（写进备份文件，便于日后核对是哪一版做的备份）</summary>
        public static string AppVersion
            => Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "";
    }
}

using FreeSql;
using NLog;
using NPOI.SS.UserModel;
using ORT一键报告.Models;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 预领（月排程 / ORT看板 / 准备领用）的数据服务：
    /// 三份数据按月归档，导入时先与库里该月的旧数据比对，把变动逐条写进「变动记录」，再整体替换该月数据；
    /// 旧月份不会被删除，任何时候都可以切回查看，也可以再次导入更新（同样会记录变动）。
    /// </summary>
    public class PreIssueService
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;
        private readonly IPermissionService _permission;

        /// <summary>一次导入最多逐条记录的变条数（超出只记摘要，避免首次导入刷出几千条）</summary>
        private const int MaxLoggedChanges = 300;

        public PreIssueService(DatabaseService db, IPermissionService permission)
        {
            _db = db;
            _permission = permission;
        }

        /* ###############################  月份与读取  ################################ */

        /// <summary>当前月份（yyyyMM）</summary>
        public static string CurrentMonth => DateTime.Now.ToString("yyyyMM", CultureInfo.InvariantCulture);

        /// <summary>
        /// 归一化月份：接受 202611 / 2026-11 / 2026/11 / 2026年11月 / DateTime / Excel 日期
        /// </summary>
        public static string NormalizeMonth(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return null;
            }
            string text = raw.Trim();
            Match m = Regex.Match(text, @"^(20\d{2})\D{0,2}(\d{1,2})");
            if (m.Success && int.TryParse(m.Groups[2].Value, out int month) && month >= 1 && month <= 12)
            {
                return m.Groups[1].Value + month.ToString("D2", CultureInfo.InvariantCulture);
            }
            if (text.Length == 6 && int.TryParse(text, out int yyyymm))
            {
                int mm = yyyymm % 100;
                return mm >= 1 && mm <= 12 ? text : null;
            }
            return null;
        }

        /// <summary>yyyyMM → 该月第一天（解析不出来返回 null）</summary>
        public static DateTime? MonthStart(string month)
        {
            string normalized = NormalizeMonth(month);
            if (normalized == null || normalized.Length != 6)
            {
                return null;
            }
            return DateTime.TryParseExact(normalized + "01", "yyyyMMdd", CultureInfo.InvariantCulture,
                DateTimeStyles.None, out DateTime date) ? date : null;
        }

        /// <summary>月份的可读文本（202611 → 2026/11），解析不出来原样返回</summary>
        public static string MonthText(string month)
        {
            string normalized = NormalizeMonth(month);
            return normalized == null || normalized.Length != 6
                ? month
                : normalized.Substring(0, 4) + "/" + normalized.Substring(4, 2);
        }

        /// <summary>
        /// 有数据的月份（三张表取并集，最新在前）；至少包含当前月
        /// </summary>
        public List<string> GetMonths()
        {
            HashSet<string> months = new(StringComparer.Ordinal) { CurrentMonth };
            try
            {
                foreach (string m in _db.FreeSql.Select<PreScheduleItem>().ToList(a => a.Month))
                {
                    if (NormalizeMonth(m) != null)
                    {
                        months.Add(NormalizeMonth(m));
                    }
                }
                foreach (string m in _db.FreeSql.Select<OrtBoardRow>().ToList(a => a.Month))
                {
                    if (NormalizeMonth(m) != null)
                    {
                        months.Add(NormalizeMonth(m));
                    }
                }
                foreach (string m in _db.FreeSql.Select<PreIssueRequisition>().ToList(a => a.Month))
                {
                    if (NormalizeMonth(m) != null)
                    {
                        months.Add(NormalizeMonth(m));
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"读取预领月份失败: {ex.Message}");
            }
            return months.OrderByDescending(m => m, StringComparer.Ordinal).ToList();
        }

        /// <summary>某月的月排程明细</summary>
        public List<PreScheduleItem> LoadSchedule(string month) =>
            _db.FreeSql.Select<PreScheduleItem>()
                .Where(a => a.Month == month)
                .OrderBy(a => a.Line).OrderBy(a => a.Day).ToList();

        /// <summary>某月的 ORT 看板</summary>
        public List<OrtBoardRow> LoadBoard(string month) =>
            _db.FreeSql.Select<OrtBoardRow>()
                .Where(a => a.Month == month)
                .OrderBy(a => a.ReportDate).ToList();

        /// <summary>某月的准备领用</summary>
        public List<PreIssueRequisition> LoadRequisition(string month) =>
            _db.FreeSql.Select<PreIssueRequisition>()
                .Where(a => a.Month == month)
                .OrderBy(a => a.State).OrderBy(a => a.PartNo).ToList();

        /// <summary>黑名单全部机种</summary>
        public List<OrtBlacklistItem> LoadBlacklist() =>
            _db.FreeSql.Select<OrtBlacklistItem>().OrderBy(a => a.Model).ToList();

        /// <summary>黑名单机种集合（比对用，忽略大小写）</summary>
        public HashSet<string> BlacklistModels()
        {
            HashSet<string> set = new(StringComparer.OrdinalIgnoreCase);
            foreach (OrtBlacklistItem item in LoadBlacklist())
            {
                if (!string.IsNullOrWhiteSpace(item.Model))
                {
                    set.Add(item.Model.Trim());
                }
            }
            return set;
        }

        /// <summary>变动记录（按月份/来源筛选，最新在前）</summary>
        public List<PreIssueChangeLog> LoadChanges(string month, string source, int limit = 800)
        {
            ISelect<PreIssueChangeLog> select = _db.FreeSql.Select<PreIssueChangeLog>();
            if (!string.IsNullOrWhiteSpace(month))
            {
                select = select.Where(a => a.Month == month);
            }
            if (!string.IsNullOrWhiteSpace(source))
            {
                select = select.Where(a => a.Source == source);
            }
            return select.OrderByDescending(a => a.Id).Limit(limit).ToList();
        }

        /// <summary>某表最近一次导入时间（界面提示「多久没更新」用）</summary>
        public DateTime? LastImportAt(string source, string month)
        {
            try
            {
                return source switch
                {
                    PreIssueCodes.SourceSchedule => _db.FreeSql.Select<PreScheduleItem>()
                        .Where(a => a.Month == month).Max(a => a.ImportedAt),
                    PreIssueCodes.SourceBoard => _db.FreeSql.Select<OrtBoardRow>()
                        .Where(a => a.Month == month).Max(a => a.ImportedAt),
                    _ => _db.FreeSql.Select<PreIssueRequisition>()
                        .Where(a => a.Month == month).Max(a => a.UpdatedAt)
                };
            }
            catch
            {
                return null;
            }
        }

        /* ###############################  月排程导入  ################################ */

        /// <summary>
        /// 导入月排程（平铺矩阵：年月 / 線別 / 1日…31日，单元格「機種/數量」）。
        /// 只做导入与更新记录，不改变表格的展示方式。
        /// </summary>
        public PreIssueImportResult ImportSchedule(string filePath)
        {
            PreIssueImportResult result = new() { File = Path.GetFileName(filePath), Source = PreIssueCodes.SourceSchedule };
            List<PreScheduleItem> parsed = ParseScheduleFile(filePath, result);
            if (parsed.Count == 0)
            {
                result.Success = false;
                result.Message = result.Message ?? "没有读到排程数据";
                return result;
            }
            // 按月份分组落库（一个文件一般只有一个月）
            foreach (IGrouping<string, PreScheduleItem> group in parsed.GroupBy(a => a.Month))
            {
                ImportScheduleMonth(group.Key, group.ToList(), result);
            }
            return result;
        }

        /// <summary>解析月排程文件（不落库）</summary>
        private List<PreScheduleItem> ParseScheduleFile(string filePath, PreIssueImportResult result)
        {
            List<PreScheduleItem> items = [];
            IWorkbook workbook = ExcelNpoi.OpenAny(filePath);
            ISheet sheet = ExcelNpoi.SheetByName(workbook, "月排程") ?? ExcelNpoi.SheetAt(workbook, 0);
            if (sheet == null)
            {
                result.Message = "文件里没有可读的工作表";
                return items;
            }
            int lastRow = ExcelNpoi.LastRow(sheet);
            int lastCol = ExcelNpoi.LastColumn(sheet);
            if (lastRow < 2 || lastCol < 3)
            {
                result.Message = "排程表看起来是空的（需要表头 + 至少一行数据）";
                return items;
            }

            // 表头：第 1 行是「年月 / 線別 / 1日 …31日」；日列取表头里的数字，取不到时按列序补齐
            Dictionary<int, int> dayColumns = [];
            for (int col = 3; col <= lastCol; col++)
            {
                string header = ExcelNpoi.CellText(sheet, 1, col);
                Match m = Regex.Match(header ?? "", @"(\d{1,2})");
                int day = m.Success ? int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture) : 0;
                if (day < 1 || day > 31)
                {
                    day = col - 2; // 表头没有日号：按第 3 列 = 1 日顺排
                }
                dayColumns[col] = day;
            }

            string fileMonth = MonthFromFileName(filePath);
            string operatorName = _permission?.CurrentUser;
            DateTime now = DateTime.Now;
            for (int row = 2; row <= lastRow; row++)
            {
                string monthCell = ExcelNpoi.CellText(sheet, row, 1);
                string lineCell = ExcelNpoi.CellText(sheet, row, 2);
                if (string.IsNullOrWhiteSpace(monthCell) && string.IsNullOrWhiteSpace(lineCell))
                {
                    continue;
                }
                string month = NormalizeMonth(monthCell) ?? fileMonth;
                if (month == null)
                {
                    if (result.Message == null)
                    {
                        result.Message = "第 1 列（年月）认不出月份，也没能从文件名推出月份（文件名里带 202611 这样的年月即可）";
                    }
                    continue;
                }
                string line = string.IsNullOrWhiteSpace(lineCell) ? "" : lineCell.Trim();
                foreach (KeyValuePair<int, int> pair in dayColumns)
                {
                    string text = ExcelNpoi.CellText(sheet, row, pair.Key);
                    if (string.IsNullOrWhiteSpace(text))
                    {
                        continue;
                    }
                    (string model, int qty) = ParseModelQty(text);
                    items.Add(new PreScheduleItem
                    {
                        Month = month,
                        Line = line,
                        Model = model,
                        Day = pair.Value,
                        Qty = qty,
                        RawText = text.Trim(),
                        SourceFile = Path.GetFileName(filePath),
                        ImportedAt = now,
                        ImportedBy = operatorName
                    });
                }
            }
            return items;
        }

        /// <summary>把「機種/數量」拆成机种与数量（数量认不出时按 0 处理，原文保留在 RawText）</summary>
        public static (string Model, int Qty) ParseModelQty(string text)
        {
            string raw = (text ?? "").Trim();
            if (raw.Length == 0)
            {
                return ("", 0);
            }
            // 数量在最后一个 / 之后（机种本身不含 /，但为稳妥仍按最后一个取）
            int slash = raw.LastIndexOf('/');
            string model = slash > 0 ? raw.Substring(0, slash).Trim() : raw;
            int qty = 0;
            if (slash > 0)
            {
                string qtyText = raw.Substring(slash + 1).Trim().Replace(",", "").Replace("　", "");
                _ = int.TryParse(qtyText, NumberStyles.Any, CultureInfo.InvariantCulture, out qty);
            }
            return (model, qty);
        }

        /// <summary>从文件名里推月份（「202611 工令排程.xlsx」→ 202611）</summary>
        private static string MonthFromFileName(string filePath)
        {
            string name = Path.GetFileNameWithoutExtension(filePath ?? "");
            Match m = Regex.Match(name, @"(20\d{2})\D{0,2}(\d{1,2})");
            if (m.Success && int.TryParse(m.Groups[2].Value, out int month) && month >= 1 && month <= 12)
            {
                return m.Groups[1].Value + month.ToString("D2", CultureInfo.InvariantCulture);
            }
            return null;
        }

        /// <summary>某月的排程导入：比对 → 记录变动 → 整体替换该月数据</summary>
        private void ImportScheduleMonth(string month, List<PreScheduleItem> newItems, PreIssueImportResult result)
        {
            List<PreScheduleItem> oldItems = _db.FreeSql.Select<PreScheduleItem>().Where(a => a.Month == month).ToList();
            Dictionary<string, PreScheduleItem> oldMap = new(StringComparer.Ordinal);
            foreach (PreScheduleItem item in oldItems)
            {
                oldMap[ScheduleKey(item)] = item;
            }
            Dictionary<string, PreScheduleItem> newMap = new(StringComparer.Ordinal);
            foreach (PreScheduleItem item in newItems)
            {
                newMap[ScheduleKey(item)] = item;
            }

            int added = 0, removed = 0, changed = 0;
            List<PreIssueChangeLog> logs = [];
            int logged = 0;
            foreach (PreScheduleItem item in newItems.Where(i => !oldMap.ContainsKey(ScheduleKey(i))))
            {
                added++;
                if (logged++ < MaxLoggedChanges)
                {
                    logs.Add(Change(month, PreIssueCodes.SourceSchedule, PreIssueCodes.KindAdded, item.Model,
                        null, null, $"{item.Line} {item.Day}日 {item.RawText}", $"来自 {item.SourceFile}"));
                }
            }
            foreach (PreScheduleItem item in oldItems.Where(i => !newMap.ContainsKey(ScheduleKey(i))))
            {
                removed++;
                if (logged++ < MaxLoggedChanges)
                {
                    logs.Add(Change(month, PreIssueCodes.SourceSchedule, PreIssueCodes.KindRemoved, item.Model,
                        null, $"{item.Line} {item.Day}日 {item.RawText}", null, "新排程里已没有这个单元格"));
                }
            }
            foreach (PreScheduleItem item in newItems)
            {
                if (!oldMap.TryGetValue(ScheduleKey(item), out PreScheduleItem old) || old.Qty == item.Qty)
                {
                    continue;
                }
                changed++;
                if (logged++ < MaxLoggedChanges)
                {
                    logs.Add(Change(month, PreIssueCodes.SourceSchedule, PreIssueCodes.KindModified, item.Model,
                        "數量", $"{old.Line} {old.Day}日 {old.Qty}", $"{item.Line} {item.Day}日 {item.Qty}",
                        $"来自 {item.SourceFile}"));
                }
            }
            if (logged >= MaxLoggedChanges)
            {
                logs.Add(Change(month, PreIssueCodes.SourceSchedule, PreIssueCodes.KindImported, "",
                    null, null, null, $"变动较多，只登记了前 {MaxLoggedChanges} 条"));
            }

            _db.FreeSql.Delete<PreScheduleItem>().Where(a => a.Month == month).ExecuteAffrows();
            _db.FreeSql.Insert(newItems).ExecuteAffrows();
            result.Added += added;
            result.Removed += removed;
            result.Changed += changed;
            result.Month ??= month;
            WriteLogs(logs, month, PreIssueCodes.SourceSchedule, added, removed, changed, result);
        }

        private static string ScheduleKey(PreScheduleItem item) =>
            $"{item.Line?.Trim()}|{item.Model?.Trim()}|{item.Day}";

        /* ###############################  ORT 看板导入  ################################ */

        /// <summary>
        /// 导入 ORT 看板原始数据：第 1 行不要、**第 2 行当表头**、数据从第 3 行起；
        /// 只取 A~O 共 15 列（P 列及之后是人工加的列，不导入）。
        /// 更新时重点比对「SIM87入庫否」：N → Y 的行就是本次可以领用的机种，记录变动并推进到准备领用。
        /// </summary>
        public PreIssueImportResult ImportBoard(string filePath)
        {
            PreIssueImportResult result = new() { File = Path.GetFileName(filePath), Source = PreIssueCodes.SourceBoard };
            List<OrtBoardRow> parsed = ParseBoardFile(filePath, result);
            if (parsed.Count == 0)
            {
                result.Success = false;
                result.Message = result.Message ?? "没有读到看板数据";
                return result;
            }
            foreach (IGrouping<string, OrtBoardRow> group in parsed.GroupBy(a => a.Month))
            {
                ImportBoardMonth(group.Key, group.ToList(), result);
            }
            return result;
        }

        private List<OrtBoardRow> ParseBoardFile(string filePath, PreIssueImportResult result)
        {
            List<OrtBoardRow> rows = [];
            IWorkbook workbook = ExcelNpoi.OpenAny(filePath);
            ISheet sheet = ExcelNpoi.SheetByName(workbook, "ORT看板") ?? ExcelNpoi.SheetAt(workbook, 0);
            if (sheet == null)
            {
                result.Message = "文件里没有可读的工作表";
                return rows;
            }
            int lastRow = ExcelNpoi.LastRow(sheet);
            int lastCol = Math.Min(ExcelNpoi.LastColumn(sheet), 15);
            if (lastRow < 2)
            {
                result.Message = "看板数据看起来是空的（原表第 2 行是表头，第 3 行起才是数据）";
                return rows;
            }

            // 表头在第 2 行（原始导出的第 1 行是不要的）；万一拿到的是已经剪掉第一行的表，第 1 行就是表头，也认
            int headerRow = FindBoardHeaderRow(sheet, lastRow, lastCol);
            if (headerRow == 0)
            {
                result.Message = "没找到看板表头（應有 日期/類別/機種/生產線別/生產日期/工令…），请确认是看板的原始导出数据";
                return rows;
            }
            Dictionary<string, int> map = [];
            for (int col = 1; col <= lastCol; col++)
            {
                string header = (ExcelNpoi.CellText(sheet, headerRow, col) ?? "").Trim();
                if (header.Length > 0 && !map.ContainsKey(header))
                {
                    map[header] = col;
                }
            }

            string operatorName = _permission?.CurrentUser;
            DateTime now = DateTime.Now;
            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                string reportDate = Text(sheet, row, map, "日期");
                string model = Text(sheet, row, map, "機種", "机种");
                string jobNo = Text(sheet, row, map, "工令");
                if (string.IsNullOrWhiteSpace(model) && string.IsNullOrWhiteSpace(jobNo))
                {
                    continue;
                }
                string month = MonthFromYmd(reportDate) ?? MonthFromYmd(Text(sheet, row, map, "生產日期", "生产日期"));
                if (month == null)
                {
                    continue;
                }
                rows.Add(new OrtBoardRow
                {
                    Month = month,
                    ReportDate = reportDate,
                    Category = Text(sheet, row, map, "類別", "类别"),
                    Model = model,
                    Line = Text(sheet, row, map, "生產線別", "生产线别"),
                    ProduceDate = Text(sheet, row, map, "生產日期", "生产日期"),
                    JobNo = jobNo,
                    Qty = Text(sheet, row, map, "生產數量", "生产数量"),
                    Hours = Text(sheet, row, map, "生產時數", "生产时数"),
                    Reported65 = Text(sheet, row, map, "65報工否", "65报工否"),
                    Stocked87 = Text(sheet, row, map, "SIM87入庫否", "SIM87入库否"),
                    ReceiveDate = Text(sheet, row, map, "收料日期"),
                    PartNo = Text(sheet, row, map, "料號", "料号"),
                    MaterialStock = Text(sheet, row, map, "材料剩餘庫存", "材料剩余库存"),
                    TestResult = Text(sheet, row, map, "測試結果", "测试结果"),
                    Forbidden94 = Text(sheet, row, map, "SIM94禁止出貨", "SIM94禁止出货"),
                    ImportedAt = now,
                    ImportedBy = operatorName
                });
            }
            return rows;
        }

        /// <summary>
        /// 找看板表头行：优先第 2 行（原始导出的写法），其次第 1 行（已经剪掉第一行的表）。
        /// 判定标准是这一行里能认出的看板列名够多（>=4 个），避免把标题行当表头。
        /// </summary>
        private static int FindBoardHeaderRow(ISheet sheet, int lastRow, int lastCol)
        {
            string[] known =
            [
                "日期", "類別", "类别", "機種", "机种", "生產線別", "生产线别", "生產日期", "生产日期",
                "工令", "生產數量", "生产数量", "生產時數", "生产时数", "65報工否", "65报工否",
                "SIM87入庫否", "SIM87入库否", "料號", "料号", "測試結果", "测试结果"
            ];
            int[] candidates = [2, 1, 3];
            foreach (int row in candidates)
            {
                if (row > lastRow)
                {
                    continue;
                }
                int hits = 0;
                bool hasModel = false;
                for (int col = 1; col <= lastCol; col++)
                {
                    string header = (ExcelNpoi.CellText(sheet, row, col) ?? "").Trim();
                    if (header.Length == 0)
                    {
                        continue;
                    }
                    if (header == "機種" || header == "机种")
                    {
                        hasModel = true;
                    }
                    if (Array.IndexOf(known, header) >= 0)
                    {
                        hits++;
                    }
                }
                if (hasModel && hits >= 4)
                {
                    return row;
                }
            }
            return 0;
        }

        private static string Text(ISheet sheet, int row, Dictionary<string, int> map, params string[] headers)
        {
            foreach (string header in headers)
            {
                if (map.TryGetValue(header, out int col))
                {
                    string text = ExcelNpoi.CellText(sheet, row, col);
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        return text.Trim();
                    }
                }
            }
            return null;
        }

        /// <summary>yyyymmdd（或 yyyy/m/d、Excel 日期）→ yyyyMM</summary>
        public static string MonthFromYmd(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }
            Match m = Regex.Match(text.Trim(), @"^(20\d{2})(\d{2})(\d{2})");
            if (m.Success)
            {
                return m.Groups[1].Value + m.Groups[2].Value;
            }
            m = Regex.Match(text.Trim(), @"^(20\d{2})\D{1,2}(\d{1,2})\D{1,2}\d{1,2}");
            if (m.Success && int.TryParse(m.Groups[2].Value, out int month) && month >= 1 && month <= 12)
            {
                return m.Groups[1].Value + month.ToString("D2", CultureInfo.InvariantCulture);
            }
            return null;
        }

        private void ImportBoardMonth(string month, List<OrtBoardRow> newRows, PreIssueImportResult result)
        {
            List<OrtBoardRow> oldRows = _db.FreeSql.Select<OrtBoardRow>().Where(a => a.Month == month).ToList();
            Dictionary<string, OrtBoardRow> oldMap = new(StringComparer.OrdinalIgnoreCase);
            foreach (OrtBoardRow row in oldRows)
            {
                oldMap[row.MatchKey] = row;
            }

            int added = 0, removed = 0, changed = 0, stocked = 0;
            List<PreIssueChangeLog> logs = [];
            HashSet<string> stockedKeys = new(StringComparer.OrdinalIgnoreCase);
            int logged = 0;

            foreach (OrtBoardRow row in newRows)
            {
                if (!oldMap.TryGetValue(row.MatchKey, out OrtBoardRow old))
                {
                    added++;
                    if (logged++ < MaxLoggedChanges)
                    {
                        logs.Add(Change(month, PreIssueCodes.SourceBoard, PreIssueCodes.KindAdded, row.Model,
                            null, null, $"{row.JobNo} {row.ReportDate} SIM87={row.Stocked87}", $"来自 {result.File}"));
                    }
                    if (PreIssueCodes.IsYes(row.Stocked87))
                    {
                        stockedKeys.Add(row.MatchKey);
                    }
                    continue;
                }
                row.HasRequisition = old.HasRequisition;
                foreach ((string Field, string OldValue, string NewValue) diff in DiffBoardRow(old, row))
                {
                    changed++;
                    if (logged++ < MaxLoggedChanges)
                    {
                        string kind = diff.Field == "SIM87入庫否" ? PreIssueCodes.KindStocked : PreIssueCodes.KindModified;
                        logs.Add(Change(month, PreIssueCodes.SourceBoard, kind, row.Model,
                            diff.Field, diff.OldValue, diff.NewValue, $"工令 {row.JobNo}"));
                    }
                    if (diff.Field == "SIM87入庫否" && PreIssueCodes.IsYes(diff.NewValue))
                    {
                        stockedKeys.Add(row.MatchKey);
                    }
                }
            }
            HashSet<string> newKeys = new(newRows.Select(r => r.MatchKey), StringComparer.OrdinalIgnoreCase);
            foreach (OrtBoardRow old in oldRows.Where(r => !newKeys.Contains(r.MatchKey)))
            {
                removed++;
                if (logged++ < MaxLoggedChanges)
                {
                    logs.Add(Change(month, PreIssueCodes.SourceBoard, PreIssueCodes.KindRemoved, old.Model,
                        null, $"{old.JobNo} {old.ReportDate}", null, "新看板里已没有这个工令"));
                }
            }
            if (logged >= MaxLoggedChanges)
            {
                logs.Add(Change(month, PreIssueCodes.SourceBoard, PreIssueCodes.KindImported, "",
                    null, null, null, $"变动较多，只登记了前 {MaxLoggedChanges} 条"));
            }

            // 入库（SIM87 = Y）的机种推进到准备领用：看板上打「已转」标记，黑名单里的不推
            HashSet<string> blacklist = BlacklistModels();
            HashSet<string> existingModels = new(
                _db.FreeSql.Select<PreIssueRequisition>().Where(a => a.Month == month).ToList()
                    .Where(a => !string.IsNullOrWhiteSpace(a.PartNo))
                    .Select(a => a.PartNo.Trim()),
                StringComparer.OrdinalIgnoreCase);

            // 第一遍：算清哪些机种这次要新建待开单（同一机种多行只建一条，工令取第一条）
            Dictionary<string, OrtBoardRow> willCreate = new(StringComparer.OrdinalIgnoreCase);
            foreach (OrtBoardRow row in newRows)
            {
                if (!PreIssueCodes.IsYes(row.Stocked87) || string.IsNullOrWhiteSpace(row.Model))
                {
                    continue;
                }
                string model = row.Model.Trim();
                if (stockedKeys.Contains(row.MatchKey))
                {
                    stocked++;
                }
                if (blacklist.Contains(model) || existingModels.Contains(model) || willCreate.ContainsKey(model))
                {
                    continue;
                }
                willCreate[model] = row;
            }
            // 看板行打「已转待开单」标记（含以前就转过的）
            foreach (OrtBoardRow row in newRows)
            {
                if (!string.IsNullOrWhiteSpace(row.Model))
                {
                    string model = row.Model.Trim();
                    row.HasRequisition = existingModels.Contains(model) || willCreate.ContainsKey(model);
                }
            }

            // 整体替换该月看板（只存 A~O 共 15 列）
            _db.FreeSql.Delete<OrtBoardRow>().Where(a => a.Month == month).ExecuteAffrows();
            _db.FreeSql.Insert(newRows).ExecuteAffrows();

            // 第二遍：真正生成待开单
            DateTime now = DateTime.Now;
            foreach (KeyValuePair<string, OrtBoardRow> pair in willCreate)
            {
                OrtBoardRow row = pair.Value;
                _db.FreeSql.Insert(new PreIssueRequisition
                {
                    Month = month,
                    PartNo = pair.Key,
                    JobNo = row.JobNo,
                    State = PreIssueCodes.PendingState,
                    CreatedAt = now,
                    CreatedBy = _permission?.CurrentUser,
                    UpdatedAt = now
                }).ExecuteAffrows();
                existingModels.Add(pair.Key);
                result.PendingCreated++;
                logs.Add(Change(month, PreIssueCodes.SourceRequisition, PreIssueCodes.KindStocked, pair.Key,
                    "SIM87入庫否", "N", "Y", $"看板入库自动转待开单（工令 {row.JobNo}）"));
            }

            result.Stocked += stocked;
            result.Added += added;
            result.Removed += removed;
            result.Changed += changed;
            result.Month ??= month;
            WriteLogs(logs, month, PreIssueCodes.SourceBoard, added, removed, changed, result);
        }

        /// <summary>看板行的逐列比对（只比 A~O 这 15 列）</summary>
        private static IEnumerable<(string Field, string OldValue, string NewValue)> DiffBoardRow(OrtBoardRow old, OrtBoardRow now)
        {
            foreach ((string field, string a, string b) in new (string, string, string)[]
            {
                ("日期", old.ReportDate, now.ReportDate),
                ("類別", old.Category, now.Category),
                ("機種", old.Model, now.Model),
                ("生產線別", old.Line, now.Line),
                ("生產日期", old.ProduceDate, now.ProduceDate),
                ("工令", old.JobNo, now.JobNo),
                ("生產數量", old.Qty, now.Qty),
                ("生產時數", old.Hours, now.Hours),
                ("65報工否", old.Reported65, now.Reported65),
                ("SIM87入庫否", old.Stocked87, now.Stocked87),
                ("收料日期", old.ReceiveDate, now.ReceiveDate),
                ("料號", old.PartNo, now.PartNo),
                ("材料剩餘庫存", old.MaterialStock, now.MaterialStock),
                ("測試結果", old.TestResult, now.TestResult),
                ("SIM94禁止出貨", old.Forbidden94, now.Forbidden94)
            })
            {
                if (!string.Equals(Norm(a), Norm(b), StringComparison.Ordinal))
                {
                    yield return (field, a, b);
                }
            }
        }

        private static string Norm(string s) => string.IsNullOrWhiteSpace(s) ? "" : s.Trim();

        /* ###############################  准备领用导入  ################################ */

        /// <summary>
        /// 导入开单数据（如 11.XLS，实际是 xlsx 格式）：只要 ORT 领用表的 10 列
        /// （VCHNO / WARE_CODE / TR_CODE / PART_NO / REQ_QTY / IVER_ID / VCH_STATUS / ISSUE_DEPT / CLOSE_CODE / PC），
        /// 多余列全部丢弃。再次导入时同步变动，VCH_STATUS、CLOSE_CODE 两列的变化会重点记录。
        /// </summary>
        public PreIssueImportResult ImportRequisition(string filePath, string defaultMonth)
        {
            PreIssueImportResult result = new() { File = Path.GetFileName(filePath), Source = PreIssueCodes.SourceRequisition };
            List<Dictionary<string, string>> rows = ParseRequisitionFile(filePath, result);
            if (rows.Count == 0)
            {
                result.Success = false;
                result.Message = result.Message ?? "没有读到开单数据";
                return result;
            }
            string operatorName = _permission?.CurrentUser;
            DateTime now = DateTime.Now;
            List<PreIssueChangeLog> logs = [];
            int logged = 0;

            // 已有记录：按单号、再按「月份 + 机种」索引
            List<PreIssueRequisition> all = _db.FreeSql.Select<PreIssueRequisition>().ToList();
            Dictionary<string, PreIssueRequisition> byVchno = new(StringComparer.OrdinalIgnoreCase);
            foreach (PreIssueRequisition r in all.Where(r => !string.IsNullOrWhiteSpace(r.Vchno)))
            {
                byVchno[r.Vchno.Trim()] = r;
            }
            // 待开单按「月份 + 机种」索引：同一个机种在不同月份各有一条待开单，不能串
            Dictionary<string, PreIssueRequisition> byPendingPart = new(StringComparer.OrdinalIgnoreCase);
            foreach (PreIssueRequisition r in all.Where(r => PreIssueCodes.IsPendingState(r.State) && !string.IsNullOrWhiteSpace(r.PartNo)))
            {
                byPendingPart[(r.Month ?? "") + "|" + r.PartNo.Trim()] = r;
            }

            foreach (Dictionary<string, string> row in rows)
            {
                string vchno = Value(row, "VCHNO");
                string partNo = Value(row, "PART_NO");
                if (string.IsNullOrWhiteSpace(vchno) && string.IsNullOrWhiteSpace(partNo))
                {
                    continue;
                }
                string month = MonthFromVchno(vchno) ?? NormalizeMonth(defaultMonth) ?? CurrentMonth;

                PreIssueRequisition target = null;
                if (!string.IsNullOrWhiteSpace(vchno) && byVchno.TryGetValue(vchno, out PreIssueRequisition hit))
                {
                    target = hit;
                }
                else
                {
                    target = FindPending(byPendingPart, month, partNo);
                }

                if (target == null)
                {
                    target = new PreIssueRequisition
                    {
                        Month = month,
                        State = PreIssueCodes.OpenedState,
                        CreatedAt = now,
                        CreatedBy = operatorName
                    };
                    ApplyRequisitionFields(target, row);
                    target.UpdatedAt = now;
                    target.UpdatedBy = operatorName;
                    _db.FreeSql.Insert(target).ExecuteAffrows();
                    result.Added++;
                    if (logged++ < MaxLoggedChanges)
                    {
                        logs.Add(Change(month, PreIssueCodes.SourceRequisition, PreIssueCodes.KindAdded, partNo,
                            null, null, vchno, "开单数据新增"));
                    }
                    if (!string.IsNullOrWhiteSpace(vchno))
                    {
                        byVchno[vchno] = target;
                    }
                    if (!string.IsNullOrWhiteSpace(partNo))
                    {
                        byPendingPart.Remove(month + "|" + partNo.Trim());
                    }
                    continue;
                }

                bool wasPending = PreIssueCodes.IsPendingState(target.State);
                foreach ((string Field, string OldValue, string NewValue) diff in DiffRequisition(target, row))
                {
                    result.Changed++;
                    if (logged++ < MaxLoggedChanges)
                    {
                        logs.Add(Change(target.Month, PreIssueCodes.SourceRequisition, PreIssueCodes.KindModified, partNo,
                            diff.Field, diff.OldValue, diff.NewValue, $"单号 {vchno}（{month}）"));
                    }
                }
                ApplyRequisitionFields(target, row);
                if (wasPending)
                {
                    target.State = PreIssueCodes.OpenedState;
                    result.Opened++;
                    logs.Add(Change(target.Month, PreIssueCodes.SourceRequisition, PreIssueCodes.KindOpened, partNo,
                        null, PreIssueCodes.PendingState, target.Vchno, "待开单已开单"));
                }
                target.UpdatedAt = now;
                target.UpdatedBy = operatorName;
                _db.FreeSql.Update<PreIssueRequisition>()
                    .Set(a => a.Vchno, target.Vchno)
                    .Set(a => a.WareCode, target.WareCode)
                    .Set(a => a.TrCode, target.TrCode)
                    .Set(a => a.PartNo, target.PartNo)
                    .Set(a => a.ReqQty, target.ReqQty)
                    .Set(a => a.IverId, target.IverId)
                    .Set(a => a.VchStatus, target.VchStatus)
                    .Set(a => a.IssueDept, target.IssueDept)
                    .Set(a => a.CloseCode, target.CloseCode)
                    .Set(a => a.Pc, target.Pc)
                    .Set(a => a.State, target.State)
                    .Set(a => a.UpdatedAt, target.UpdatedAt)
                    .Set(a => a.UpdatedBy, target.UpdatedBy)
                    .Where(a => a.Id == target.Id)
                    .ExecuteAffrows();
                if (!string.IsNullOrWhiteSpace(target.Vchno))
                {
                    byVchno[target.Vchno.Trim()] = target;
                }
                if (!string.IsNullOrWhiteSpace(target.PartNo))
                {
                    byPendingPart.Remove((target.Month ?? "") + "|" + target.PartNo.Trim());
                }
            }

            if (logged >= MaxLoggedChanges)
            {
                logs.Add(Change(null, PreIssueCodes.SourceRequisition, PreIssueCodes.KindImported, "",
                    null, null, null, $"变动较多，只登记了前 {MaxLoggedChanges} 条"));
            }
            result.Month ??= NormalizeMonth(defaultMonth) ?? MonthFromVchno(rows.Select(r => Value(r, "VCHNO")).FirstOrDefault());
            WriteLogs(logs, result.Month, PreIssueCodes.SourceRequisition, result.Added, 0, result.Changed, result);
            result.Month = NormalizeMonth(defaultMonth) ?? MonthFromVchno(rows.Select(r => Value(r, "VCHNO")).FirstOrDefault());
            return result;
        }

        /// <summary>单号里带年月：T01**2611**00166 → 202611（前缀字母 + 2 位部门 + YY + MM + 5 位流水）</summary>
        public static string MonthFromVchno(string vchno)
        {
            if (string.IsNullOrWhiteSpace(vchno))
            {
                return null;
            }
            Match m = Regex.Match(vchno.Trim(), @"^[A-Za-z]{1,3}(\d{2})(\d{2})(\d{2})\d{3,}$");
            if (m.Success && int.TryParse(m.Groups[3].Value, out int month) && month >= 1 && month <= 12)
            {
                return "20" + m.Groups[2].Value + month.ToString("D2", CultureInfo.InvariantCulture);
            }
            return null;
        }

        private List<Dictionary<string, string>> ParseRequisitionFile(string filePath, PreIssueImportResult result)
        {
            List<Dictionary<string, string>> rows = [];
            string[] wanted =
            [
                "VCHNO", "WARE_CODE", "TR_CODE", "PART_NO", "REQ_QTY",
                "IVER_ID", "VCH_STATUS", "ISSUE_DEPT", "CLOSE_CODE", "PC"
            ];
            IWorkbook workbook = ExcelNpoi.OpenAny(filePath);
            for (int index = 0; index < workbook.NumberOfSheets; index++)
            {
                ISheet sheet = ExcelNpoi.SheetAt(workbook, index);
                if (sheet == null)
                {
                    continue;
                }
                int headerRow = FindRequisitionHeader(sheet);
                if (headerRow <= 0)
                {
                    continue;
                }
                Dictionary<string, int> map = [];
                int lastCol = ExcelNpoi.LastColumn(sheet);
                for (int col = 1; col <= lastCol; col++)
                {
                    string header = (ExcelNpoi.CellText(sheet, headerRow, col) ?? "").Trim().ToUpperInvariant();
                    if (wanted.Contains(header) && !map.ContainsKey(header))
                    {
                        map[header] = col;
                    }
                }
                if (!map.ContainsKey("VCHNO"))
                {
                    continue;
                }
                int lastRow = ExcelNpoi.LastRow(sheet);
                for (int row = headerRow + 1; row <= lastRow; row++)
                {
                    Dictionary<string, string> item = [];
                    foreach (string key in wanted)
                    {
                        item[key] = map.TryGetValue(key, out int col) ? Clean(ExcelNpoi.CellText(sheet, row, col)) : null;
                    }
                    if (string.IsNullOrWhiteSpace(item["VCHNO"]) && string.IsNullOrWhiteSpace(item["PART_NO"]))
                    {
                        continue;
                    }
                    rows.Add(item);
                }
                break;
            }
            if (rows.Count == 0)
            {
                result.Message = "没找到开单数据的表头（需要 VCHNO / PART_NO / VCH_STATUS / CLOSE_CODE 等列）";
            }
            return rows;
        }

        private static int FindRequisitionHeader(ISheet sheet)
        {
            int lastRow = Math.Min(ExcelNpoi.LastRow(sheet), 10);
            for (int row = 1; row <= lastRow; row++)
            {
                int lastCol = Math.Min(ExcelNpoi.LastColumn(sheet), 40);
                for (int col = 1; col <= lastCol; col++)
                {
                    if (string.Equals((ExcelNpoi.CellText(sheet, row, col) ?? "").Trim(), "VCHNO", StringComparison.OrdinalIgnoreCase))
                    {
                        return row;
                    }
                }
            }
            return 0;
        }

        /// <summary>
        /// 在待开单里找同一个机种：先按「月份 + 机种」精确找，找不到再退一步按机种找（跨月的开单数据）。
        /// </summary>
        private static PreIssueRequisition FindPending(Dictionary<string, PreIssueRequisition> index, string month, string partNo)
        {
            if (index.Count == 0 || string.IsNullOrWhiteSpace(partNo))
            {
                return null;
            }
            string part = partNo.Trim();
            if (index.TryGetValue((month ?? "") + "|" + part, out PreIssueRequisition exact))
            {
                return exact;
            }
            foreach (KeyValuePair<string, PreIssueRequisition> pair in index)
            {
                if (pair.Key.EndsWith("|" + part, StringComparison.OrdinalIgnoreCase))
                {
                    return pair.Value;
                }
            }
            return null;
        }

        private static string Value(Dictionary<string, string> row, string key) =>
            row.TryGetValue(key, out string value) ? value : null;

        private static string Clean(string text) => string.IsNullOrWhiteSpace(text) ? null : text.Trim();

        private static IEnumerable<(string Field, string OldValue, string NewValue)> DiffRequisition(
            PreIssueRequisition old, Dictionary<string, string> row)
        {
            (string Field, string OldValue, string NewValue)[] pairs =
            [
                ("VCHNO", old.Vchno, Value(row, "VCHNO")),
                ("WARE_CODE", old.WareCode, Value(row, "WARE_CODE")),
                ("TR_CODE", old.TrCode, Value(row, "TR_CODE")),
                ("PART_NO", old.PartNo, Value(row, "PART_NO")),
                ("REQ_QTY", old.ReqQty, Value(row, "REQ_QTY")),
                ("IVER_ID", old.IverId, Value(row, "IVER_ID")),
                ("VCH_STATUS", old.VchStatus, Value(row, "VCH_STATUS")),
                ("ISSUE_DEPT", old.IssueDept, Value(row, "ISSUE_DEPT")),
                ("CLOSE_CODE", old.CloseCode, Value(row, "CLOSE_CODE")),
                ("PC", old.Pc, Value(row, "PC"))
            ];
            foreach ((string Field, string OldValue, string NewValue) pair in pairs)
            {
                if (!string.Equals(Norm(pair.OldValue), Norm(pair.NewValue), StringComparison.Ordinal))
                {
                    yield return pair;
                }
            }
        }

        private static void ApplyRequisitionFields(PreIssueRequisition target, Dictionary<string, string> row)
        {
            target.Vchno = Value(row, "VCHNO");
            target.WareCode = Value(row, "WARE_CODE");
            target.TrCode = Value(row, "TR_CODE");
            target.PartNo = Value(row, "PART_NO");
            target.ReqQty = Value(row, "REQ_QTY");
            target.IverId = Value(row, "IVER_ID");
            target.VchStatus = Value(row, "VCH_STATUS");
            target.IssueDept = Value(row, "ISSUE_DEPT");
            target.CloseCode = Value(row, "CLOSE_CODE");
            target.Pc = Value(row, "PC");
        }

        /* ###############################  准备领用的手工/看板补全  ################################ */

        /// <summary>
        /// 把看板里「SIM87入庫否 = Y 且不在黑名单、还没进准备领用」的机种补成待开单。
        /// 返回新建条数。
        /// </summary>
        public int SyncPendingFromBoard(string month, out List<string> created)
        {
            created = [];
            HashSet<string> blacklist = BlacklistModels();
            HashSet<string> existing = new(
                _db.FreeSql.Select<PreIssueRequisition>().Where(a => a.Month == month).ToList()
                    .Where(a => !string.IsNullOrWhiteSpace(a.PartNo)).Select(a => a.PartNo.Trim()),
                StringComparer.OrdinalIgnoreCase);
            List<OrtBoardRow> board = LoadBoard(month);
            DateTime now = DateTime.Now;
            foreach (IGrouping<string, OrtBoardRow> group in board
                .Where(b => PreIssueCodes.IsYes(b.Stocked87) && !string.IsNullOrWhiteSpace(b.Model))
                .GroupBy(b => b.Model.Trim(), StringComparer.OrdinalIgnoreCase))
            {
                string model = group.Key;
                if (blacklist.Contains(model) || existing.Contains(model))
                {
                    continue;
                }
                OrtBoardRow first = group.First();
                _db.FreeSql.Insert(new PreIssueRequisition
                {
                    Month = month,
                    PartNo = model,
                    JobNo = first.JobNo,
                    State = PreIssueCodes.PendingState,
                    CreatedAt = now,
                    CreatedBy = _permission?.CurrentUser,
                    UpdatedAt = now
                }).ExecuteAffrows();
                existing.Add(model);
                created.Add(model);
            }
            if (created.Count > 0)
            {
                _db.FreeSql.Insert(new PreIssueChangeLog
                {
                    Month = month,
                    Source = PreIssueCodes.SourceRequisition,
                    Kind = PreIssueCodes.KindStocked,
                    Target = string.Join("、", created.Take(20)),
                    Note = $"从看板补全待开单 {created.Count} 个机种",
                    Operator = _permission?.CurrentUser,
                    CreatedAt = now
                }).ExecuteAffrows();
            }
            return created.Count;
        }

        /// <summary>手工加一条待开单</summary>
        public bool AddPending(string month, string model, string jobNo, out string error)
        {
            error = null;
            if (string.IsNullOrWhiteSpace(model))
            {
                error = "机种不能为空";
                return false;
            }
            string m = model.Trim();
            if (BlacklistModels().Contains(m))
            {
                error = $"「{m}」在黑名单里（ORT 做不了，不需要领用）";
                return false;
            }
            bool exists = _db.FreeSql.Select<PreIssueRequisition>()
                .Where(a => a.Month == month && a.PartNo == m).Count() > 0;
            if (exists)
            {
                error = $"「{m}」在这个月已经有一条记录了";
                return false;
            }
            DateTime now = DateTime.Now;
            _db.FreeSql.Insert(new PreIssueRequisition
            {
                Month = month,
                PartNo = m,
                JobNo = string.IsNullOrWhiteSpace(jobNo) ? null : jobNo.Trim(),
                State = PreIssueCodes.PendingState,
                CreatedAt = now,
                CreatedBy = _permission?.CurrentUser,
                UpdatedAt = now
            }).ExecuteAffrows();
            _db.FreeSql.Insert(new PreIssueChangeLog
            {
                Month = month,
                Source = PreIssueCodes.SourceRequisition,
                Kind = PreIssueCodes.KindAdded,
                Target = m,
                Note = "手工添加待开单",
                Operator = _permission?.CurrentUser,
                CreatedAt = now
            }).ExecuteAffrows();
            return true;
        }

        /// <summary>删除一条待开单（已开单的不让删，避免与开单数据对不上）</summary>
        public bool DeletePending(long id, out string error)
        {
            error = null;
            PreIssueRequisition row = _db.FreeSql.Select<PreIssueRequisition>().Where(a => a.Id == id).First();
            if (row == null)
            {
                error = "记录已不存在";
                return false;
            }
            if (!PreIssueCodes.IsPendingState(row.State) || !string.IsNullOrWhiteSpace(row.Vchno))
            {
                error = "只有待开单的记录可以删除";
                return false;
            }
            _db.FreeSql.Delete<PreIssueRequisition>().Where(a => a.Id == id).ExecuteAffrows();
            _db.FreeSql.Insert(new PreIssueChangeLog
            {
                Month = row.Month,
                Source = PreIssueCodes.SourceRequisition,
                Kind = PreIssueCodes.KindRemoved,
                Target = row.PartNo,
                Note = "删除待开单",
                Operator = _permission?.CurrentUser,
                CreatedAt = DateTime.Now
            }).ExecuteAffrows();
            return true;
        }

        /* ###############################  黑名单  ################################ */

        /// <summary>加入黑名单（重复加入直接忽略）</summary>
        public bool AddToBlacklist(string model, string reason, out string error)
        {
            error = null;
            if (string.IsNullOrWhiteSpace(model))
            {
                error = "机种不能为空";
                return false;
            }
            string m = model.Trim();
            if (_db.FreeSql.Select<OrtBlacklistItem>().Where(a => a.Model == m).Count() > 0)
            {
                error = $"「{m}」已经在黑名单里了";
                return false;
            }
            _db.FreeSql.Insert(new OrtBlacklistItem
            {
                Model = m,
                Reason = string.IsNullOrWhiteSpace(reason) ? null : reason.Trim(),
                CreatedBy = _permission?.CurrentUser,
                CreatedAt = DateTime.Now
            }).ExecuteAffrows();
            _db.FreeSql.Insert(new PreIssueChangeLog
            {
                Source = PreIssueCodes.SourceBlacklist,
                Kind = PreIssueCodes.KindBlacklist,
                Target = m,
                NewValue = "加入黑名单",
                Note = reason,
                Operator = _permission?.CurrentUser,
                CreatedAt = DateTime.Now
            }).ExecuteAffrows();
            return true;
        }

        /// <summary>移出黑名单</summary>
        public bool RemoveFromBlacklist(string model, out string error)
        {
            error = null;
            if (string.IsNullOrWhiteSpace(model))
            {
                error = "机种不能为空";
                return false;
            }
            string m = model.Trim();
            int deleted = _db.FreeSql.Delete<OrtBlacklistItem>().Where(a => a.Model == m).ExecuteAffrows();
            if (deleted == 0)
            {
                error = $"「{m}」不在黑名单里";
                return false;
            }
            _db.FreeSql.Insert(new PreIssueChangeLog
            {
                Source = PreIssueCodes.SourceBlacklist,
                Kind = PreIssueCodes.KindBlacklist,
                Target = m,
                OldValue = "加入黑名单",
                NewValue = "移出黑名单",
                Operator = _permission?.CurrentUser,
                CreatedAt = DateTime.Now
            }).ExecuteAffrows();
            return true;
        }

        /* ###############################  变动记录  ################################ */

        private PreIssueChangeLog Change(string month, string source, string kind, string target,
            string field, string oldValue, string newValue, string note) => new()
        {
            Month = month,
            Source = source,
            Kind = kind,
            Target = string.IsNullOrWhiteSpace(target) ? null : target.Trim(),
            Field = field,
            OldValue = Truncate(oldValue),
            NewValue = Truncate(newValue),
            Note = Truncate(note),
            Operator = _permission?.CurrentUser,
            CreatedAt = DateTime.Now
        };

        private static string Truncate(string text) =>
            text == null ? null : (text.Length > 500 ? text.Substring(0, 500) + "…" : text);

        /// <summary>写变动记录并补齐导入摘要（每次导入都至少留一条）</summary>
        private void WriteLogs(List<PreIssueChangeLog> logs, string month, string source,
            int added, int removed, int changed, PreIssueImportResult result)
        {
            try
            {
                if (logs.Count > 0)
                {
                    _db.FreeSql.Insert(logs).ExecuteAffrows();
                }
                string summary = $"导入 {result.File}：新增 {added}、删除 {removed}、修改 {changed}"
                    + (result.PendingCreated > 0 ? $"、转待开单 {result.PendingCreated}" : "");
                _db.FreeSql.Insert(new PreIssueChangeLog
                {
                    Month = month,
                    Source = source,
                    Kind = PreIssueCodes.KindImported,
                    Target = result.File,
                    Note = summary,
                    Operator = _permission?.CurrentUser,
                    CreatedAt = DateTime.Now
                }).ExecuteAffrows();
                _logger.Info($"预领导入 {source} {summary}");
            }
            catch (Exception ex)
            {
                _logger.Warn($"写预领变动记录失败: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// 一次预领导入的结果（界面提示与变动记录用）
    /// </summary>
    public class PreIssueImportResult
    {
        /// <summary>来源表（月排程 / ORT看板 / 准备领用）</summary>
        public string Source { get; set; }

        /// <summary>导入的文件名</summary>
        public string File { get; set; }

        /// <summary>涉及月份</summary>
        public string Month { get; set; }

        /// <summary>是否成功</summary>
        public bool Success { get; set; } = true;

        /// <summary>提示信息（失败原因 / 额外说明）</summary>
        public string Message { get; set; }

        /// <summary>新增条数</summary>
        public int Added { get; set; }

        /// <summary>删除条数</summary>
        public int Removed { get; set; }

        /// <summary>修改的列数</summary>
        public int Changed { get; set; }

        /// <summary>本次发现入库（SIM87 = Y）的行数</summary>
        public int Stocked { get; set; }

        /// <summary>自动转到准备领用的待开单条数</summary>
        public int PendingCreated { get; set; }

        /// <summary>准备领用：本次由待开单补全成已开单的条数</summary>
        public int Opened { get; set; }

        /// <summary>界面提示文本</summary>
        public string Summary()
        {
            if (!Success)
            {
                return string.IsNullOrWhiteSpace(Message) ? "导入失败" : Message;
            }
            List<string> parts = [];
            if (Added > 0) parts.Add($"新增 {Added}");
            if (Removed > 0) parts.Add($"删除 {Removed}");
            if (Changed > 0) parts.Add($"修改 {Changed}");
            if (Opened > 0) parts.Add($"补全开单 {Opened}");
            if (PendingCreated > 0) parts.Add($"转待开单 {PendingCreated}");
            string detail = parts.Count == 0 ? "没有变动" : string.Join("、", parts);
            return string.IsNullOrWhiteSpace(Message) ? detail : detail + "（" + Message + "）";
        }
    }
}

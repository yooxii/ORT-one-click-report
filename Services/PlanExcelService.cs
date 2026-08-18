using NLog;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.OleObject;
using ORT一键报告.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 计划管理 Excel 导入导出服务：
    /// - 以计划表为基底、领用表辅助，两表数据按工令/机种名称等关联合并到同一行；
    /// - 导入"成品領用記錄(领用表)"与"ORT Test Schedule(计划表)"，按唯一键合并重复项；
    /// - 提取领用表 S/N 列的嵌入 OLE 对象保存到 Data\OleFiles，命名为 {简短日期}_{领用单据号}_{机种名称}；
    /// - 从 plans 表重新导出为领用表/计划表。
    /// </summary>
    public class PlanExcelService
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;
        private readonly IPermissionService _permission;

        private static readonly string[] RequisitionHeaders =
        [
            "領用/日期", "領料單据號", "機種名稱", "測試項目", "領出/數量", "S/N", "D/C",
            "REV.", "Work Order", "回綫 RT 工令", "回線/數量", "線別", "回線/日期",
            "入庫退料/單据號", "入庫/數量", "入庫日期", "备注"
        ];

        private static readonly string[] ScheduleHeaders =
        [
            "工作編號/Job No", "產品別/Product", "客戶別/Customer", "機種名/Part No", "階 段/Stage",
            "測試項目/Test Item", "樣品數/Sample Size", "試驗時間/Test Period", "負責人/Owner",
            "開始日期/Start Date", "結束日期/End Date", "完成狀況/Status", "上傳系統/Upload e-lab", "備 考/Remark"
        ];

        public PlanExcelService(DatabaseService db, IPermissionService permission)
        {
            _db = db;
            _permission = permission;
        }

        static PlanExcelService()
        {
            // EPPlus 8 非商业许可，确保任何入口使用本服务前许可已设置
            ExcelPackage.License.SetNonCommercialPersonal("Lucas");
        }

        /* ###############################  导入  ################################ */

        /// <summary>
        /// 导入领用表（成品領用記錄），与已有记录合并到同一行：
        /// 依次按 领料单据号 → 回线RT工令 → WorkOrder → 机种+测试项目+数量 匹配
        /// </summary>
        /// <returns>(新增数, 更新数)</returns>
        public (int added, int updated) ImportRequisition(string filePath)
        {
            _logger.Info($"导入领用表: {filePath}");
            int added = 0, updated = 0;
            int year = ParseYearFromFileName(filePath);
            using FileStream fs = new(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using ExcelPackage package = new(fs);
            ExcelWorksheet ws = package.Workbook.Worksheets[0];
            (int headerRow, Dictionary<string, int> map) = FindHeaderRow(ws, "領料單据號");
            if (headerRow == 0)
            {
                throw new InvalidDataException("未找到领用表表头(領料單据號)");
            }

            // 按行收集 S/N 列的 OLE 对象（表头原文为"S/N"）
            int snCol = map["S/N"];
            Dictionary<int, ExcelOleObject> oleByRow = [];
            foreach (var drawing in ws.Drawings)
            {
                if (drawing is ExcelOleObject ole && ole.From.Column + 1 == snCol)
                {
                    oleByRow[ole.From.Row + 1] = ole;
                }
            }

            int endRow = ws.Dimension?.End.Row ?? 0;
            for (int r = headerRow + 1; r <= endRow; r++)
            {
                string requisitionNo = Cell(ws, r, map, "領料單据號");
                string modelName = Cell(ws, r, map, "機種名稱");
                string requisitionDate = Cell(ws, r, map, "領用日期");
                string returnRt = Cell(ws, r, map, "回綫RT工令");
                string workOrder = Cell(ws, r, map, "WorkOrder");
                string testItem = Cell(ws, r, map, "測試項目");
                string outQty = Cell(ws, r, map, "領出數量");
                // 整行关键字段均为空则跳过
                if (requisitionNo == null && modelName == null && workOrder == null)
                {
                    continue;
                }

                // 合并到同一行：依次按唯一键、工令、同机种+领用日期与开始日期一致 匹配已有记录（含已导入的计划行）
                Plan existing = FindExistingRow(requisitionNo, returnRt, workOrder, modelName, plan_dateValue: ParseChineseDate(requisitionDate, year));
                Plan plan = existing ?? new Plan { CreatedBy = _permission.CurrentUser, CreatedAt = DateTime.Now };

                plan.RequisitionDate = requisitionDate;
                plan.RequisitionDateValue = ParseChineseDate(requisitionDate, year);
                plan.RequisitionNo = requisitionNo;
                plan.ModelName = modelName;
                plan.TestItem = testItem;
                plan.OutQty = outQty;
                plan.SN = Cell(ws, r, map, "S/N");
                plan.DC = Cell(ws, r, map, "D/C");
                plan.Rev = Cell(ws, r, map, "REV");
                plan.WorkOrder = workOrder;
                plan.ReturnRtOrder = returnRt;
                plan.ReturnQty = Cell(ws, r, map, "回線數量");
                plan.LineNo = Cell(ws, r, map, "線別");
                plan.ReturnDate = Cell(ws, r, map, "回線日期");
                plan.ReturnDateValue = ParseChineseDate(plan.ReturnDate, year);
                plan.StockInNo = Cell(ws, r, map, "入庫退料單据號");
                plan.StockInQty = Cell(ws, r, map, "入庫數量");
                plan.StockInDate = Cell(ws, r, map, "入庫日期");
                plan.StockInDateValue = ParseChineseDate(plan.StockInDate, year);
                // 备注：不覆盖计划表已有的备注（计划表备注信息更丰富），仅在为空时补充
                string remark = Cell(ws, r, map, "备注");
                if (remark != null || existing == null)
                {
                    plan.Remark = remark ?? plan.Remark;
                }

                // 该行存在嵌入的 OLE 对象（SN清单文件）时提取保存
                if (oleByRow.TryGetValue(r, out ExcelOleObject ole))
                {
                    string fileName = SaveOleObject(ole, filePath, GetShortDate(requisitionDate), requisitionNo, modelName);
                    if (fileName != null)
                    {
                        plan.SnFilePath = fileName;
                        _logger.Info($"行{r}的OLE对象已提取保存: {fileName}");
                    }
                }

                plan.UpdatedBy = _permission.CurrentUser;
                plan.UpdatedAt = DateTime.Now;

                if (existing == null)
                {
                    _db.FreeSql.Insert(plan).ExecuteAffrows();
                    added++;
                }
                else
                {
                    _db.FreeSql.Update<Plan>().SetSource(plan).Where(p => p.Id == plan.Id).ExecuteAffrows();
                    updated++;
                }
            }
            _logger.Info($"领用表导入完成: 新增{added}条, 更新{updated}条");
            return (added, updated);
        }

        /// <summary>
        /// 导入计划表（ORT Test Schedule）：以计划表为基底，
        /// 依次按 工作編號 → 备注中的工令 → 机种+测试项目 匹配已有行（含领用表导入的行）合并。
        /// 备注无工令且工作編號非 Q 开头的行无法关联领用数据，收入 unmatched 列表供提示。
        /// </summary>
        /// <returns>(新增数, 更新数, 未匹配到领用数据的工作編號列表)</returns>
        public (int added, int updated, List<string> unmatched) ImportSchedule(string filePath)
        {
            _logger.Info($"导入计划表: {filePath}");
            int added = 0, updated = 0;
            List<string> unmatched = [];
            using FileStream fs = new(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using ExcelPackage package = new(fs);

            // 优先选择名为 Schedule 的工作表，否则取第一个
            ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault(s => s.Name == "Schedule")
                ?? package.Workbook.Worksheets[0];
            (int headerRow, Dictionary<string, int> map) = FindHeaderRow(ws, "工作編號");
            if (headerRow == 0)
            {
                throw new InvalidDataException("未找到计划表表头(工作編號)");
            }

            int endRow = ws.Dimension?.End.Row ?? 0;
            for (int r = headerRow + 1; r <= endRow; r++)
            {
                string jobNo = Cell(ws, r, map, "工作編號");
                if (jobNo == null)
                {
                    continue; // 无工作編號的行视为统计/空行，跳过
                }

                string modelName = Cell(ws, r, map, "機種名");
                string testItem = Cell(ws, r, map, "測試項目");
                string remark = Cell(ws, r, map, "Remark");
                DateTime? startDateValue = ParseAnyDate(Cell(ws, r, map, "開始日期"));

                Plan existing = _db.FreeSql.Select<Plan>().Where(p => p.JobNo == jobNo).First();
                if (existing == null)
                {
                    existing = FindScheduleMatch(modelName, testItem, remark, startDateValue);
                }
                // 未能合并到领用数据且工作編號非 Q 开头（Q开头为可靠性试验计划，通常无领用记录）：提示用户
                if (existing == null && !jobNo.StartsWith("Q", StringComparison.OrdinalIgnoreCase))
                {
                    unmatched.Add($"{jobNo} ({modelName ?? "-"})");
                }
                Plan plan = existing ?? new Plan { CreatedBy = _permission.CurrentUser, CreatedAt = DateTime.Now };

                plan.JobNo = jobNo;
                plan.Product = Cell(ws, r, map, "產品別");
                plan.Customer = Cell(ws, r, map, "客戶別");
                plan.ModelName = modelName;
                plan.Stage = Cell(ws, r, map, "階段");
                plan.TestItem = testItem;
                plan.SampleSize = Cell(ws, r, map, "樣品數");
                plan.TestPeriod = Cell(ws, r, map, "試驗時間");
                plan.Owner = Cell(ws, r, map, "負責人");
                plan.StartDate = Cell(ws, r, map, "開始日期");
                plan.StartDateValue = startDateValue;
                plan.EndDate = Cell(ws, r, map, "結束日期");
                plan.EndDateValue = ParseAnyDate(plan.EndDate);
                plan.Status = Cell(ws, r, map, "完成狀況");
                plan.UploadELab = Cell(ws, r, map, "上傳系統");
                // 备注：计划表备注信息更丰富，优先使用；为空时保留原备注
                plan.Remark = remark ?? plan.Remark;
                plan.UpdatedBy = _permission.CurrentUser;
                plan.UpdatedAt = DateTime.Now;

                if (existing == null)
                {
                    _db.FreeSql.Insert(plan).ExecuteAffrows();
                    added++;
                }
                else
                {
                    _db.FreeSql.Update<Plan>().SetSource(plan).Where(p => p.Id == plan.Id).ExecuteAffrows();
                    updated++;
                }
            }
            _logger.Info($"计划表导入完成: 新增{added}条, 更新{updated}条, 未匹配{unmatched.Count}条");
            return (added, updated, unmatched);
        }

        /* ###############################  导出  ################################ */

        /// <summary>
        /// 导出为领用表（成品領退管理表格式），并将已提取的SN文件以OLE对象嵌回S/N列
        /// </summary>
        public void ExportRequisition(string savePath)
        {
            _logger.Info($"导出领用表: {savePath}");
            List<Plan> plans = _db.FreeSql.Select<Plan>()
                .Where(p => p.RequisitionNo != null)
                .OrderBy(p => p.Id)
                .ToList();

            using ExcelPackage package = new();
            ExcelWorksheet ws = package.Workbook.Worksheets.Add("退管理表");
            ws.Cells[1, 2].Value = "ORT 課試驗成品領退管理表";
            ws.Cells[3, 2].LoadFromArrays(new object[][] { RequisitionHeaders });

            int r = 4;
            foreach (Plan plan in plans)
            {
                ws.Cells[r, 2].Value = plan.RequisitionDate;
                ws.Cells[r, 3].Value = plan.RequisitionNo;
                ws.Cells[r, 4].Value = plan.ModelName;
                ws.Cells[r, 5].Value = plan.TestItem;
                ws.Cells[r, 6].Value = plan.OutQty;
                ws.Cells[r, 7].Value = plan.SN;
                ws.Cells[r, 8].Value = plan.DC;
                ws.Cells[r, 9].Value = plan.Rev;
                ws.Cells[r, 10].Value = plan.WorkOrder;
                ws.Cells[r, 11].Value = plan.ReturnRtOrder;
                ws.Cells[r, 12].Value = plan.ReturnQty;
                ws.Cells[r, 13].Value = plan.LineNo;
                ws.Cells[r, 14].Value = plan.ReturnDate;
                ws.Cells[r, 15].Value = plan.StockInNo;
                ws.Cells[r, 16].Value = plan.StockInQty;
                ws.Cells[r, 17].Value = plan.StockInDate;
                ws.Cells[r, 18].Value = plan.Remark;

                // SN文件存在时以OLE对象形式嵌回S/N列，尽量还原原表形态
                string snFile = _db.ResolveAttachmentPath(plan.SnFilePath);
                if (!string.IsNullOrWhiteSpace(plan.SnFilePath) && File.Exists(snFile))
                {
                    Utils.Report.EmbedOleObjectWithEpplus(ws, snFile, $"G{r}");
                }
                r++;
            }
            package.SaveAs(new FileInfo(savePath));
            _logger.Info($"领用表导出完成，共{plans.Count}条");
        }

        /// <summary>
        /// 导出为计划表（ORT Test Schedule格式）
        /// </summary>
        public void ExportSchedule(string savePath)
        {
            _logger.Info($"导出计划表: {savePath}");
            List<Plan> plans = _db.FreeSql.Select<Plan>()
                .Where(p => p.JobNo != null)
                .OrderBy(p => p.Id)
                .ToList();

            using ExcelPackage package = new();
            ExcelWorksheet ws = package.Workbook.Worksheets.Add("Schedule");
            ws.Cells[1, 3].Value = "ORT Test Schedule";
            ws.Cells[3, 2].LoadFromArrays(new object[][] { ScheduleHeaders });

            int r = 4;
            foreach (Plan plan in plans)
            {
                ws.Cells[r, 2].Value = plan.JobNo;
                ws.Cells[r, 3].Value = plan.Product;
                ws.Cells[r, 4].Value = plan.Customer;
                ws.Cells[r, 5].Value = plan.ModelName;
                ws.Cells[r, 6].Value = plan.Stage;
                ws.Cells[r, 7].Value = plan.TestItem;
                ws.Cells[r, 8].Value = plan.SampleSize;
                ws.Cells[r, 9].Value = plan.TestPeriod;
                ws.Cells[r, 10].Value = plan.Owner;
                ws.Cells[r, 11].Value = plan.StartDate;
                ws.Cells[r, 12].Value = plan.EndDate;
                ws.Cells[r, 13].Value = plan.Status;
                ws.Cells[r, 14].Value = plan.UploadELab;
                ws.Cells[r, 15].Value = plan.Remark;
                r++;
            }
            package.SaveAs(new FileInfo(savePath));
            _logger.Info($"计划表导出完成，共{plans.Count}条");
        }

        /// <summary>
        /// 清空全部计划数据（数据库文件保留，表结构不变）
        /// </summary>
        /// <returns>删除的记录数</returns>
        public int ClearAll()
        {
            int n = _db.FreeSql.Delete<Plan>().Where("1=1").ExecuteAffrows();
            _logger.Info($"已清空全部计划数据，共{n}条");
            return n;
        }

        /* ###############################  功能函数  ################################ */

        /// <summary>
        /// 领用表行与已有记录的匹配：领料单据号 → 回线RT工令 → WorkOrder（含已有行）→ 已导入计划行的备注中包含该工令
        /// → 无工令时兑底：同机种且领用日期与开始日期一致（视为同一记录）。
        /// 不用机种+数量兑底匹配，避免把不同领用单合并导致工令丢失。
        /// </summary>
        private Plan FindExistingRow(string requisitionNo, string returnRt, string workOrder, string modelName, DateTime? plan_dateValue)
        {
            if (requisitionNo != null)
            {
                Plan p = _db.FreeSql.Select<Plan>().Where(x => x.RequisitionNo == requisitionNo).First();
                if (p != null)
                {
                    return p;
                }
            }
            if (returnRt != null)
            {
                Plan p = _db.FreeSql.Select<Plan>().Where(x => x.ReturnRtOrder == returnRt).First();
                if (p != null)
                {
                    return p;
                }
            }
            if (workOrder != null)
            {
                Plan p = _db.FreeSql.Select<Plan>().Where(x => x.WorkOrder == workOrder).First();
                if (p != null)
                {
                    return p;
                }
                // 先导入计划表的场景：计划行备注中包含该工令且尚未合并领用数据
                List<Plan> scheduleRows = _db.FreeSql.Select<Plan>()
                    .Where(x => x.JobNo != null && x.WorkOrder == null && x.Remark != null)
                    .ToList();
                foreach (Plan s in scheduleRows)
                {
                    if (WorkOrderMatch(s.Remark, workOrder))
                    {
                        return s;
                    }
                }
            }
            // 无工令时的兑底判定：同机种且领用日期与开始日期一致
            if (modelName != null && plan_dateValue != null)
            {
                List<Plan> candidates = _db.FreeSql.Select<Plan>()
                    .Where(x => x.JobNo != null && x.RequisitionNo == null && x.ModelName == modelName && x.StartDateValue != null)
                    .ToList();
                Plan byDate = candidates.FirstOrDefault(c => c.StartDateValue.Value.Date == plan_dateValue.Value.Date);
                if (byDate != null)
                {
                    return byDate;
                }
            }
            return null;
        }

        /// <summary>
        /// 计划表行与已有记录（含领用表导入的行）的匹配：
        /// 优先用备注中的工令关联（双向包含匹配，工令唯一足以定位，不限机种）；
        /// 无工令时按 同机种+领用日期与开始日期一致 判定；最后机种+测试项目兑底（仅限尚无工作編號的行）。
        /// </summary>
        private Plan FindScheduleMatch(string modelName, string testItem, string remark, DateTime? startDateValue)
        {
            if (remark != null)
            {
                List<Plan> candidates = _db.FreeSql.Select<Plan>()
                    .Where(p => p.JobNo == null && p.WorkOrder != null)
                    .ToList();
                foreach (Plan c in candidates)
                {
                    if (WorkOrderMatch(remark, c.WorkOrder))
                    {
                        return c;
                    }
                }
            }
            // 无工令时的判定：同机种且领用日期与开始日期一致，视为同一记录
            if (modelName != null && startDateValue != null)
            {
                List<Plan> candidates = _db.FreeSql.Select<Plan>()
                    .Where(p => p.JobNo == null && p.ModelName == modelName && p.RequisitionDateValue != null)
                    .ToList();
                Plan byDate = candidates.FirstOrDefault(c => c.RequisitionDateValue.Value.Date == startDateValue.Value.Date);
                if (byDate != null)
                {
                    return byDate;
                }
            }
            if (modelName != null && testItem != null)
            {
                return _db.FreeSql.Select<Plan>()
                    .Where(p => p.JobNo == null && p.ModelName == modelName && p.TestItem == testItem)
                    .First();
            }
            return null;
        }

        /// <summary>
        /// 工令匹配：相等、备注包含工令、或备注本身即工令（≥12位字母数字）且被工令包含（备注可能被截断）
        /// </summary>
        private static bool WorkOrderMatch(string remark, string workOrder)
            => remark.Contains(workOrder)
            || (remark.Length >= 12 && remark.All(char.IsLetterOrDigit) && workOrder.Contains(remark));

        /// <summary>
        /// 从导入文件名中提取年份（如 "_2026.成品領用記錄" -> 2026），无则用当前年份
        /// </summary>
        private static int ParseYearFromFileName(string filePath)
        {
            Match m = Regex.Match(Path.GetFileName(filePath), @"(19|20)\d{2}");
            return m.Success ? int.Parse(m.Value) : DateTime.Now.Year;
        }

        /// <summary>
        /// 解析中文日期文本（如 "1月9日"）为指定年份的 DateTime；失败时尝试通用解析
        /// </summary>
        private static DateTime? ParseChineseDate(string text, int year)
        {
            if (text == null)
            {
                return null;
            }
            Match m = Regex.Match(text, @"(\d{1,2})\s*月\s*(\d{1,2})\s*日");
            if (m.Success
                && int.TryParse(m.Groups[1].Value, out int month)
                && int.TryParse(m.Groups[2].Value, out int day))
            {
                try
                {
                    return new DateTime(year, month, day);
                }
                catch
                {
                    return null;
                }
            }
            return ParseAnyDate(text);
        }

        /// <summary>
        /// 通用日期解析（如 "2026/1/2"），失败返回null
        /// </summary>
        private static DateTime? ParseAnyDate(string text)
        {
            return DateTime.TryParse(text, out DateTime dt) ? dt : null;
        }

        /// <summary>
        /// 在前10行内寻找包含指定关键字的表头行，返回(表头行号, 规范化表头文本->列号)映射；未找到返回(0, null)
        /// </summary>
        private static (int, Dictionary<string, int>) FindHeaderRow(ExcelWorksheet ws, string headerKey)
        {
            int endRow = Math.Min(ws.Dimension?.End.Row ?? 0, 10);
            int endCol = ws.Dimension?.End.Column ?? 0;
            for (int r = 1; r <= endRow; r++)
            {
                Dictionary<string, int> map = [];
                bool hit = false;
                for (int c = 1; c <= endCol; c++)
                {
                    string key = Norm(ws.Cells[r, c].Text);
                    if (key == "")
                    {
                        continue;
                    }
                    if (key.Contains(Norm(headerKey)))
                    {
                        hit = true;
                    }
                    map[key] = c;
                }
                if (hit)
                {
                    return (r, map);
                }
            }
            return (0, null);
        }

        /// <summary>
        /// 按表头关键字(包含匹配, 忽略空白)读取单元格文本；空白返回null。
        /// 注：原表头中的"/"实为换行显示，规范化后不含斜杠，搜索关键字也不要带斜杠。
        /// </summary>
        private static string Cell(ExcelWorksheet ws, int row, Dictionary<string, int> map, string headerKey)
        {
            string normKey = Norm(headerKey);
            foreach (KeyValuePair<string, int> kv in map)
            {
                if (kv.Key.Contains(normKey))
                {
                    return NullIfEmpty(ws.Cells[row, kv.Value].Text);
                }
            }
            return null;
        }

        /// <summary>
        /// 去除空白字符（仅限 Unicode 空白，不用 \s：.NET Framework 下 \s 可能把 '/' 也当作空白匹配），用于表头/关键字匹配
        /// </summary>
        private static string Norm(string s) => Regex.Replace(s ?? "", "[\\p{Z}\\p{C}\\t\\r\\n ]", "");

        private static string NullIfEmpty(string s) => string.IsNullOrWhiteSpace(s) ? null : s.Trim();

        /// <summary>
        /// 从日期文本中提取"月日"4位简短日期，如 "1月9日" -> "0109"；解析失败使用当前日期
        /// </summary>
        private static string GetShortDate(string dateText)
        {
            Match m = Regex.Match(dateText ?? "", @"(\d{1,2})\s*月\s*(\d{1,2})\s*日");
            if (m.Success)
            {
                return $"{int.Parse(m.Groups[1].Value):D2}{int.Parse(m.Groups[2].Value):D2}";
            }
            if (DateTime.TryParse(dateText, out DateTime dt))
            {
                return dt.ToString("MMdd");
            }
            return DateTime.Now.ToString("MMdd");
        }

        /// <summary>
        /// 提取 OLE 嵌入对象保存到附件目录，命名 {简短日期}_{领用单据号}_{机种名称}.ext。
        /// 优先直接解析 xlsx 包内 embeddings 数据（稳定无异常）；失败时再用 EPPlus API 回退。
        /// </summary>
        private string SaveOleObject(ExcelOleObject ole, string sourceFilePath, string shortDate, string requisitionNo, string modelName)
        {
            try
            {
                // 优先 zip 直读：对 ProgId="工作表" 等对象 EPPlus 的 GetEmbeddedObjectBytes 会抛异常（触发调试器一级异常中断），避开之
                byte[] bytes = ExtractOleBytesFromZip(sourceFilePath, ole.Name);
                if (bytes == null || bytes.Length == 0)
                {
                    try
                    {
                        bytes = ole.GetEmbeddedObjectBytes();
                    }
                    catch (Exception ex)
                    {
                        _logger.Warn($"GetEmbeddedObjectBytes({ole.Name})失败: {ex.Message}");
                    }
                }
                if (bytes == null || bytes.Length == 0)
                {
                    _logger.Warn($"OLE对象({ole.Name})无嵌入数据，跳过");
                    return null;
                }
                // 优先按文件头判断真实类型，其次按ProgId推断
                string ext = GetExtensionByBytes(bytes) ?? GetExtensionByProgId(ole.ProgId);
                string baseName = $"{shortDate}_{CleanFileName(requisitionNo ?? "无单据号")}_{CleanFileName(modelName ?? "无机种名")}";
                string fileName = baseName + ext;
                string fullPath = Path.Combine(_db.OleDir, fileName);
                // 同名文件已存在且内容相同则直接复用
                if (File.Exists(fullPath))
                {
                    if (bytes.SequenceEqual(File.ReadAllBytes(fullPath)))
                    {
                        return fileName;
                    }
                    fullPath = Path.Combine(_db.OleDir, $"{baseName}_{DateTime.Now:HHmmss}{ext}");
                    fileName = Path.GetFileName(fullPath);
                }
                File.WriteAllBytes(fullPath, bytes);
                return fileName;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"提取OLE对象({ole.Name})失败");
                return null;
            }
        }

        /// <summary>
        /// 回退方案：直接解析 xlsx 包，按 OLE 对象名称找到对应 drawings 关系，读取 embeddings 的 bin 字节
        /// </summary>
        private byte[] ExtractOleBytesFromZip(string xlsxPath, string oleName)
        {
            try
            {
                using ZipArchive zip = ZipFile.OpenRead(xlsxPath);
                foreach (ZipArchiveEntry drawingEntry in zip.Entries
                    .Where(e => e.FullName.StartsWith("xl/drawings/") && e.FullName.EndsWith(".xml")))
                {
                    XDocument doc;
                    using (Stream s = drawingEntry.Open())
                    {
                        doc = XDocument.Load(s);
                    }
                    XNamespace xdr = doc.Root.Name.Namespace;
                    XElement oleNode = doc.Descendants(xdr + "oleObject")
                        .FirstOrDefault(n => (string)n.Attribute("name") == oleName);
                    if (oleNode == null)
                    {
                        continue;
                    }
                    string rId = oleNode.Attributes()
                        .FirstOrDefault(a => a.Name.LocalName == "id" && a.Name.NamespaceName.Contains("relationships"))?.Value;
                    if (rId == null)
                    {
                        continue;
                    }
                    string relsPath = "xl/drawings/_rels/" + Path.GetFileName(drawingEntry.FullName) + ".rels";
                    ZipArchiveEntry relsEntry = zip.GetEntry(relsPath);
                    if (relsEntry == null)
                    {
                        continue;
                    }
                    XDocument rels;
                    using (Stream s2 = relsEntry.Open())
                    {
                        rels = XDocument.Load(s2);
                    }
                    XElement rel = rels.Descendants().FirstOrDefault(n => (string)n.Attribute("Id") == rId);
                    string target = (string)rel?.Attribute("Target");
                    if (string.IsNullOrEmpty(target))
                    {
                        continue;
                    }
                    string binPath = target.StartsWith("../")
                        ? "xl/" + target.Substring(3)
                        : "xl/drawings/" + target;
                    ZipArchiveEntry binEntry = zip.GetEntry(binPath);
                    if (binEntry == null)
                    {
                        continue;
                    }
                    using Stream s3 = binEntry.Open();
                    using MemoryStream ms = new();
                    s3.CopyTo(ms);
                    return ms.ToArray();
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"zip回退提取OLE({oleName})失败: {ex.Message}");
            }
            return null;
        }

        /// <summary>
        /// 根据嵌入数据的文件头判断真实文件类型；无法识别返回null
        /// </summary>
        private static string GetExtensionByBytes(byte[] bytes)
        {
            if (bytes.Length >= 4 && bytes[0] == 0x50 && bytes[1] == 0x4B) return ".xlsx"; // zip容器(xlsx/docx/pptx等)
            if (bytes.Length >= 4 && bytes[0] == 0xD0 && bytes[1] == 0xCF && bytes[2] == 0x11 && bytes[3] == 0xE0) return ".xls"; // OLE复合文档
            if (bytes.Length >= 5 && bytes[0] == 0x25 && bytes[1] == 0x50 && bytes[2] == 0x44 && bytes[3] == 0x46) return ".pdf";
            return null;
        }

        /// <summary>
        /// 根据OLE对象的ProgId推断原始文件扩展名
        /// </summary>
        private static string GetExtensionByProgId(string progId)
        {
            string id = progId?.ToLower() ?? "";
            if (id.Contains("工作表") || id.Contains("worksheet")) return ".xls";
            if (id.Contains("excel.sheet.12") || id.Contains("xlsm") || id.Contains("csv")) return ".xlsx";
            if (id.Contains("excel.sheet.8") || id.Contains("excel.sheet")) return ".xls";
            if (id.Contains("word.document.12")) return ".docx";
            if (id.Contains("word.document")) return ".doc";
            if (id.Contains("powerpoint")) return ".pptx";
            if (id.Contains("pdffile") || id.Contains("acrobat")) return ".pdf";
            if (id.Contains("packager") || id.Contains("package")) return ".dat";
            return ".xlsx";
        }

        /// <summary>
        /// 清理文件名中的非法字符
        /// </summary>
        private static string CleanFileName(string name)
        {
            string cleaned = Regex.Replace(name ?? "", $"[{Regex.Escape(new string(Path.GetInvalidFileNameChars()))}]", "_").Trim();
            return cleaned == "" ? "_" : cleaned;
        }
    }
}

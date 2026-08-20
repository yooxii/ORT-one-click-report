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
    /// 计划管理 Excel 导入导出服务（领退表与计划表分表存储）：
    /// - 导入"成品領用記錄(领退表)" → requisitions 表；导入"ORT Test Schedule(计划表)" → plans 表；
    /// - 提取领退表 S/N 列的嵌入 OLE 对象保存到 Data\OleFiles；
    /// - 从两表分别重新导出为领退表/计划表。
    /// </summary>
    public class PlanExcelService
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;
        private readonly IPermissionService _permission;

        private static readonly string[] RequisitionHeaders =
        [
            "領用/日期", "領料單据號", "機種名稱", "領出/數量", "S/N", "D/C",
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
        /// 导入领退表（成品領用記錄）到 requisitions 表，按 领料单据号/回线RT工令/WorkOrder 合并
        /// </summary>
        /// <returns>(新增数, 更新数)</returns>
        public (int added, int updated) ImportRequisition(string filePath)
        {
            _logger.Info($"导入领退表: {filePath}");
            int added = 0, updated = 0;
            int year = ParseYearFromFileName(filePath);
            using FileStream fs = new(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using ExcelPackage package = new(fs);
            ExcelWorksheet ws = package.Workbook.Worksheets[0];
            (int headerRow, Dictionary<string, int> map) = FindHeaderRow(ws, "領料單据號");
            if (headerRow == 0)
            {
                throw new InvalidDataException("未找到领退表表头(領料單据號)");
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
                string workOrder = Cell(ws, r, map, "WorkOrder");
                // 整行关键字段均为空则跳过
                if (requisitionNo == null && modelName == null && workOrder == null)
                {
                    continue;
                }

                // 按领料单据号匹配合并；未命中时按 WorkOrder 更新最早一条（WorkOrder 可重复）
                Requisition existing = requisitionNo != null
                    ? _db.FreeSql.Select<Requisition>().Where(x => x.RequisitionNo == requisitionNo).First()
                    : null;
                if (existing == null && workOrder != null)
                {
                    existing = _db.FreeSql.Select<Requisition>().Where(x => x.WorkOrder == workOrder).First();
                }
                Requisition plan = existing ?? new Requisition { CreatedBy = _permission.CurrentUser, CreatedAt = DateTime.Now };

                // 已存在且该行未提供领料单据号时，用已有单据号（保持主键稳定）
                if (requisitionNo == null && existing != null)
                {
                    requisitionNo = existing.RequisitionNo;
                }

                plan.RequisitionDate = ParseAnyDate(Cell(ws, r, map, "領用日期"), year);
                plan.RequisitionNo = requisitionNo;
                plan.ModelName = modelName;
                plan.OutQty = Cell(ws, r, map, "領出數量");
                plan.SN = Cell(ws, r, map, "S/N");
                plan.DC = Cell(ws, r, map, "D/C");
                plan.Rev = Cell(ws, r, map, "REV");
                plan.WorkOrder = workOrder;
                plan.ReturnRtOrder = Cell(ws, r, map, "回綫RT工令");
                plan.ReturnQty = Cell(ws, r, map, "回線數量");
                plan.LineNo = Cell(ws, r, map, "線別");
                plan.ReturnDate = ParseAnyDate(Cell(ws, r, map, "回線日期"), year);
                plan.StockInNo = Cell(ws, r, map, "入庫退料單据號");
                plan.StockInQty = Cell(ws, r, map, "入庫數量");
                plan.StockInDate = ParseAnyDate(Cell(ws, r, map, "入庫日期"), year);
                plan.Remark = Cell(ws, r, map, "备注");

                // 该行存在嵌入的 OLE 对象（SN清单文件）时提取保存
                if (oleByRow.TryGetValue(r, out ExcelOleObject ole))
                {
                    string fileName = SaveOleObject(ole, filePath, GetShortDate(plan.RequisitionDate), requisitionNo, modelName);
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
                    _db.FreeSql.Update<Requisition>().SetSource(plan).Where(p => p.Id == plan.Id).ExecuteAffrows();
                    updated++;
                }
            }
            _logger.Info($"领退表导入完成: 新增{added}条, 更新{updated}条");
            return (added, updated);
        }

        /// <summary>
        /// 导入计划表（ORT Test Schedule）到 plans 表，按 工作編號 合并
        /// </summary>
        /// <returns>(新增数, 更新数, 未匹配到领用数据的工作編號列表)</returns>
        public (int added, int updated, List<string> unmatched) ImportSchedule(string filePath)
        {
            _logger.Info($"导入计划表: {filePath}");
            int added = 0, updated = 0;
            int year = ParseYearFromFileName(filePath);
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

                Plan existing = _db.FreeSql.Select<Plan>().Where(p => p.JobNo == jobNo).First();
                Plan plan = existing ?? new Plan { CreatedBy = _permission.CurrentUser, CreatedAt = DateTime.Now };

                plan.JobNo = jobNo;
                plan.Product = Cell(ws, r, map, "產品別");
                plan.Customer = Cell(ws, r, map, "客戶別");
                plan.ModelName = Cell(ws, r, map, "機種名");
                plan.Stage = Cell(ws, r, map, "階段");
                plan.TestItem = Cell(ws, r, map, "測試項目");
                plan.SampleSize = Cell(ws, r, map, "樣品數");
                plan.TestPeriod = Cell(ws, r, map, "試驗時間");
                plan.Owner = Cell(ws, r, map, "負責人");
                plan.StartDate = ParseAnyDate(Cell(ws, r, map, "開始日期"), year);
                plan.EndDate = ParseAnyDate(Cell(ws, r, map, "結束日期"), year);
                plan.Status = Cell(ws, r, map, "完成狀況");
                plan.UploadELab = Cell(ws, r, map, "上傳系統");
                plan.Remark = Cell(ws, r, map, "Remark");
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
            _logger.Info($"计划表导入完成: 新增{added}条, 更新{updated}条");
            return (added, updated, unmatched);
        }

        /* ###############################  导出  ################################ */

        /// <summary>
        /// 导出为领退表（成品領退管理表格式），并将已提取的SN文件以OLE对象嵌回S/N列
        /// </summary>
        public void ExportRequisition(string savePath)
        {
            _logger.Info($"导出领退表: {savePath}");
            List<Requisition> plans = _db.FreeSql.Select<Requisition>()
                .Where(p => p.RequisitionNo != null)
                .OrderBy(p => p.Id)
                .ToList();

            using ExcelPackage package = new();
            ExcelWorksheet ws = package.Workbook.Worksheets.Add("退管理表");
            ws.Cells[1, 2].Value = "ORT 課試驗成品領退管理表";
            ws.Cells[3, 2].LoadFromArrays(new object[][] { RequisitionHeaders });

            int r = 4;
            foreach (Requisition plan in plans)
            {
                ws.Cells[r, 2].Value = plan.RequisitionDate;
                ws.Cells[r, 3].Value = plan.RequisitionNo;
                ws.Cells[r, 4].Value = plan.ModelName;
                ws.Cells[r, 5].Value = plan.OutQty;
                ws.Cells[r, 6].Value = plan.SN;
                ws.Cells[r, 7].Value = plan.DC;
                ws.Cells[r, 8].Value = plan.Rev;
                ws.Cells[r, 9].Value = plan.WorkOrder;
                ws.Cells[r, 10].Value = plan.ReturnRtOrder;
                ws.Cells[r, 11].Value = plan.ReturnQty;
                ws.Cells[r, 12].Value = plan.LineNo;
                ws.Cells[r, 13].Value = plan.ReturnDate;
                ws.Cells[r, 14].Value = plan.StockInNo;
                ws.Cells[r, 15].Value = plan.StockInQty;
                ws.Cells[r, 16].Value = plan.StockInDate;
                ws.Cells[r, 17].Value = plan.Remark;

                // SN文件存在时以OLE对象形式嵌回S/N列，尽量还原原表形态
                string snFile = _db.ResolveAttachmentPath(plan.SnFilePath);
                if (!string.IsNullOrWhiteSpace(plan.SnFilePath) && File.Exists(snFile))
                {
                    Utils.Report.EmbedOleObjectWithEpplus(ws, snFile, $"F{r}");
                }
                r++;
            }
            package.SaveAs(new FileInfo(savePath));
            _logger.Info($"领退表导出完成，共{plans.Count}条");
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
            int m = _db.FreeSql.Delete<Requisition>().Where("1=1").ExecuteAffrows();
            _logger.Info($"已清空全部计划数据，计划表{n}条，领退表{m}条");
            return n + m;
        }

        /* ###############################  自动编号  ################################ */

        /// <summary>
        /// 生成回线RT工令：RTAH{当前年月}{编号}，编号为当月第多少个回线工令（两位数字）
        /// </summary>
        public string GenerateReturnRtOrder(DateTime date)
        {
            string ym = date.ToString("yyMM");
            int count = (int)_db.FreeSql.Select<Requisition>()
                .Where(r => r.ReturnRtOrder != null && r.ReturnRtOrder.StartsWith("RTAH" + ym))
                .Count();
            return $"RTAH{ym}{count + 1:D2}";
        }

        /// <summary>
        /// 生成工作编号：{prefix}{当前年月}{编号}，编号为当月第多少个工作编号（两位数字）
        /// </summary>
        public string GenerateJobNo(DateTime date, string prefix)
        {
            string ym = date.ToString("yyMM");
            int count = (int)_db.FreeSql.Select<Plan>()
                .Where(p => p.JobNo != null && p.JobNo.StartsWith(prefix + ym))
                .Count();
            return $"{prefix}{ym}{count + 1:D2}";
        }

        /* ###############################  功能函数  ################################ */

        /// <summary>
        /// 从前10行内寻找包含指定关键字的表头行，返回(表头行号, 规范化表头文本->列号)映射；未找到返回(0, null)
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
        private static string GetShortDate(DateTime? date)
        {
            if (date != null)
            {
                return date.Value.ToString("MMdd");
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

        /// <summary>
        /// 从导入文件名中提取年份（如 "_2026.成品領用記錄" -> 2026），无则用当前年份
        /// </summary>
        private static int ParseYearFromFileName(string filePath)
        {
            Match m = Regex.Match(Path.GetFileName(filePath), @"(19|20)\d{2}");
            return m.Success ? int.Parse(m.Value) : DateTime.Now.Year;
        }

        /// <summary>
        /// 通用日期解析：支持 "2026/8/18"、"2026-8-18"、"8月7日"（年份推断）等格式，失败返回null
        /// </summary>
        private static DateTime? ParseAnyDate(string text, int? fallbackYear = null)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }
            // 1. 中文格式：月日（年份推断）
            Match m = Regex.Match(text, @"(\d{1,2})\s*月\s*(\d{1,2})\s*日");
            if (m.Success
                && int.TryParse(m.Groups[1].Value, out int month)
                && int.TryParse(m.Groups[2].Value, out int day))
            {
                int year = fallbackYear ?? DateTime.Now.Year;
                try
                {
                    return new DateTime(year, month, day);
                }
                catch
                {
                    return null;
                }
            }
            // 2. 数字格式：2026/8/18、2026-8-18、8/18 等（不变文化环境）
            string normalized = text.Trim().Replace('-', '/');
            if (DateTime.TryParseExact(normalized,
                    ["yyyy/M/d", "yyyy/MM/dd", "M/d", "M/d/yyyy", "yyyy/M", "yyyy/M/d H:mm", "yyyy/M/d HH:mm"],
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None,
                    out DateTime parsed))
            {
                return parsed;
            }
            // 3. 兼容回退
            return DateTime.TryParse(text, out DateTime dt) ? dt : null;
        }
    }
}

using NLog;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using ORT一键报告.Models;
using ORT一键报告.Utils;
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
    /// - 提取领退表 S/N 列的嵌入 OLE 对象保存到 Data\OleFiles（zip 直读，不经 Excel）；
    /// - 从两表分别重新导出为领退表/计划表（NPOI 写入；SN 附件用 Excel COM 嵌回）。
    /// </summary>
    public class PlanExcelService
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;
        private readonly IPermissionService _permission;

        // 表头文案与原表一致：单元格里的"/"实为换行显示（用 \n 写出来），
        // 这样导入端按"去掉空白后包含关键字"匹配时仍能命中（写成字面 "/" 会导致
        // 領用/日期 这类关键字中间夹了斜杠而匹配不到，重新导入时会丢掉日期）
        private static readonly string[] RequisitionHeaders =
        [
            "領用\n日期", "領料單据號", "機種名稱", "領出\n數量", "S/N", "D/C",
            "REV.", "Work Order", "回綫 RT 工令", "回線\n數量", "線別", "回線\n日期",
            "入庫退料\n單据號", "入庫\n數量", "入庫日期", "备注"
        ];

        private static readonly string[] ScheduleHeaders =
        [
            "工作編號\nJob No", "產品別\nProduct", "客戶別\nCustomer", "機種名\nPart No", "階 段\nStage",
            "測試項目\nTest Item", "樣品數\nSample Size", "試驗時間\nTest Period", "負責人\nOwner",
            "開始日期\nStart Date", "結束日期\nEnd Date", "完成狀況\nStatus", "上傳系統\nUpload e-lab", "備 考\nRemark"
        ];

        /// <summary>
        /// 最近一次导出的 OLE 附件嵌入结果（导出后由界面读取，用于如实提示"附件是否真的写进去了"）
        /// </summary>
        public OleEmbedResult LastEmbedResult { get; private set; }

        public PlanExcelService(DatabaseService db, IPermissionService permission)
        {
            _db = db;
            _permission = permission;
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
            XSSFWorkbook wb = ExcelNpoi.OpenRead(filePath);
            try
            {
                ISheet ws = ExcelNpoi.SheetAt(wb, 0);
                (int headerRow, Dictionary<string, int> map) = FindHeaderRow(ws, "領料單据號");
                if (headerRow == 0)
                {
                    throw new InvalidDataException("未找到领退表表头(領料單据號)");
                }

                // 按行收集 S/N 列的 OLE 对象（表头原文为"S/N"）：zip 直读，不依赖任何 Excel 库
                int snCol = map["S/N"];
                Dictionary<int, (string Name, string ProgId, byte[] Bytes)> oleByRow = [];
                foreach ((int Row, int Col, string Name, string ProgId, byte[] Bytes) ole in ExtractOleObjectsFromZip(filePath))
                {
                    if (ole.Col == snCol && ole.Bytes != null && ole.Bytes.Length > 0)
                    {
                        oleByRow[ole.Row] = (ole.Name, ole.ProgId, ole.Bytes);
                    }
                }

                int endRow = ExcelNpoi.LastRow(ws);
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
                    if (oleByRow.TryGetValue(r, out (string Name, string ProgId, byte[] Bytes) ole))
                    {
                        string fileName = SaveOleObject(ole.Name, ole.ProgId, ole.Bytes, GetShortDate(plan.RequisitionDate), requisitionNo, modelName);
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
            }
            finally
            {
                wb.Close();
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
            XSSFWorkbook wb = ExcelNpoi.OpenRead(filePath);
            try
            {
                // 优先选择名为 Schedule 的工作表，否则取第一个
                ISheet ws = ExcelNpoi.SheetByName(wb, "Schedule") ?? ExcelNpoi.SheetAt(wb, 0);
                (int headerRow, Dictionary<string, int> map) = FindHeaderRow(ws, "工作編號");
                if (headerRow == 0)
                {
                    throw new InvalidDataException("未找到计划表表头(工作編號)");
                }

                int endRow = ExcelNpoi.LastRow(ws);
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
            }
            finally
            {
                wb.Close();
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

            List<OleEmbedRequest> oleRequests = [];
            XSSFWorkbook wb = ExcelNpoi.Create();
            try
            {
                ISheet ws = wb.CreateSheet("退管理表");
                ExcelNpoi.SetCell(ws, 1, 2, "ORT 課試驗成品領退管理表");
                WriteHeaderRow(ws, 3, 2, RequisitionHeaders);

                int r = 4;
                foreach (Requisition plan in plans)
                {
                    ExcelNpoi.SetCell(ws, r, 2, plan.RequisitionDate);
                    ExcelNpoi.SetCell(ws, r, 3, plan.RequisitionNo);
                    ExcelNpoi.SetCell(ws, r, 4, plan.ModelName);
                    ExcelNpoi.SetCell(ws, r, 5, plan.OutQty);
                    ExcelNpoi.SetCell(ws, r, 6, plan.SN);
                    ExcelNpoi.SetCell(ws, r, 7, plan.DC);
                    ExcelNpoi.SetCell(ws, r, 8, plan.Rev);
                    ExcelNpoi.SetCell(ws, r, 9, plan.WorkOrder);
                    ExcelNpoi.SetCell(ws, r, 10, plan.ReturnRtOrder);
                    ExcelNpoi.SetCell(ws, r, 11, plan.ReturnQty);
                    ExcelNpoi.SetCell(ws, r, 12, plan.LineNo);
                    ExcelNpoi.SetCell(ws, r, 13, plan.ReturnDate);
                    ExcelNpoi.SetCell(ws, r, 14, plan.StockInNo);
                    ExcelNpoi.SetCell(ws, r, 15, plan.StockInQty);
                    ExcelNpoi.SetCell(ws, r, 16, plan.StockInDate);
                    ExcelNpoi.SetCell(ws, r, 17, plan.Remark);

                    // SN文件存在时以OLE对象形式嵌回S/N列，尽量还原原表形态
                    // （NPOI 只负责写数据，OLE 嵌入在保存后由 Excel COM 统一完成）
                    string snFile = _db.ResolveAttachmentPath(plan.SnFilePath);
                    if (!string.IsNullOrWhiteSpace(plan.SnFilePath) && File.Exists(snFile))
                    {
                        oleRequests.Add(new OleEmbedRequest
                        {
                            ObjectPath = snFile,
                            SheetName = "退管理表",
                            TopLeftAddress = $"F{r}",
                            WidthPx = 100,
                            HeightPx = 100,
                            OffsetXPx = 10,
                            OffsetYPx = 10
                        });
                    }
                    r++;
                }
                ExcelNpoi.Save(wb, savePath);
            }
            finally
            {
                wb.Close();
            }
            LastEmbedResult = ExcelOleEmbedder.Embed(savePath, oleRequests);
            _logger.Info($"领退表导出完成，共{plans.Count}条");
        }

        /// <summary>
        /// 导出为计划表（ORT Test Schedule格式）
        /// </summary>
        public void ExportSchedule(string savePath)
        {
            _logger.Info($"导出计划表: {savePath}");
            LastEmbedResult = null; // 计划表不嵌附件
            List<Plan> plans = _db.FreeSql.Select<Plan>()
                .Where(p => p.JobNo != null)
                .OrderBy(p => p.Id)
                .ToList();

            XSSFWorkbook wb = ExcelNpoi.Create();
            try
            {
                ISheet ws = wb.CreateSheet("Schedule");
                ExcelNpoi.SetCell(ws, 1, 3, "ORT Test Schedule");
                WriteHeaderRow(ws, 3, 2, ScheduleHeaders);

                int r = 4;
                foreach (Plan plan in plans)
                {
                    ExcelNpoi.SetCell(ws, r, 2, plan.JobNo);
                    ExcelNpoi.SetCell(ws, r, 3, plan.Product);
                    ExcelNpoi.SetCell(ws, r, 4, plan.Customer);
                    ExcelNpoi.SetCell(ws, r, 5, plan.ModelName);
                    ExcelNpoi.SetCell(ws, r, 6, plan.Stage);
                    ExcelNpoi.SetCell(ws, r, 7, plan.TestItem);
                    ExcelNpoi.SetCell(ws, r, 8, plan.SampleSize);
                    ExcelNpoi.SetCell(ws, r, 9, plan.TestPeriod);
                    ExcelNpoi.SetCell(ws, r, 10, plan.Owner);
                    ExcelNpoi.SetCell(ws, r, 11, plan.StartDate);
                    ExcelNpoi.SetCell(ws, r, 12, plan.EndDate);
                    ExcelNpoi.SetCell(ws, r, 13, plan.Status);
                    ExcelNpoi.SetCell(ws, r, 14, plan.UploadELab);
                    ExcelNpoi.SetCell(ws, r, 15, plan.Remark);
                    r++;
                }
                ExcelNpoi.Save(wb, savePath);
            }
            finally
            {
                wb.Close();
            }
            _logger.Info($"计划表导出完成，共{plans.Count}条");
        }

        /// <summary>
        /// 写表头行（原 EPPlus 的 LoadFromArrays）
        /// </summary>
        private static void WriteHeaderRow(ISheet ws, int row, int startCol, string[] headers)
        {
            for (int i = 0; i < headers.Length; i++)
            {
                ExcelNpoi.SetCell(ws, row, startCol + i, headers[i]);
            }
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
        private static (int, Dictionary<string, int>) FindHeaderRow(ISheet ws, string headerKey)
        {
            int endRow = Math.Min(ExcelNpoi.LastRow(ws), 10);
            int endCol = ExcelNpoi.LastColumn(ws);
            for (int r = 1; r <= endRow; r++)
            {
                Dictionary<string, int> map = [];
                bool hit = false;
                for (int c = 1; c <= endCol; c++)
                {
                    string key = Norm(ExcelNpoi.CellText(ws, r, c));
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
        private static string Cell(ISheet ws, int row, Dictionary<string, int> map, string headerKey)
        {
            string normKey = Norm(headerKey);
            foreach (KeyValuePair<string, int> kv in map)
            {
                if (kv.Key.Contains(normKey))
                {
                    return NullIfEmpty(ExcelNpoi.CellText(ws, row, kv.Value));
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
        /// 保存已提取的 OLE 嵌入对象到附件目录，命名 {简短日期}_{领用单据号}_{机种名称}.ext。
        /// 数据由 zip 直读得到（不经 Excel，也不依赖任何 Excel 库）。
        /// </summary>
        private string SaveOleObject(string oleName, string progId, byte[] bytes, string shortDate, string requisitionNo, string modelName)
        {
            try
            {
                if (bytes == null || bytes.Length == 0)
                {
                    _logger.Warn($"OLE对象({oleName})无嵌入数据，跳过");
                    return null;
                }
                // 优先按文件头判断真实类型，其次按ProgId推断
                string ext = GetExtensionByBytes(bytes) ?? GetExtensionByProgId(progId);
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
                _logger.Error(ex, $"提取OLE对象({oleName})失败");
                return null;
            }
        }

        /// <summary>
        /// 直接解析 xlsx 包，读出全部嵌入的 OLE 对象：左上角位置（1 基行列）、名称、ProgId、embeddings 字节。
        /// 兼容两种存放方式：
        /// 1) 旧式（Excel 常见）：xl/worksheets/sheetN.xml 的 &lt;oleObjects&gt; 里，锚点在同元素的 &lt;anchor&gt; 内；
        /// 2) 新式：xl/drawings/drawingN.xml 的 &lt;xdr:oleObject&gt; 里，锚点在祖先 twoCellAnchor 的 &lt;xdr:from&gt; 内。
        /// </summary>
        private List<(int Row, int Col, string Name, string ProgId, byte[] Bytes)> ExtractOleObjectsFromZip(string xlsxPath)
        {
            List<(int, int, string, string, byte[])> result = [];
            try
            {
                // 用 FileStream + ZipArchive（不用 ZipFile 静态方法）：避免对 System.IO.Compression.FileSystem 的依赖
                using FileStream stream = new(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using ZipArchive zip = new(stream, ZipArchiveMode.Read);

                // 1) 旧式：工作表内的 oleObjects
                foreach (ZipArchiveEntry sheetEntry in zip.Entries
                    .Where(e => e.FullName.StartsWith("xl/worksheets/") && e.FullName.EndsWith(".xml")
                             && !e.FullName.Contains("/_rels/")))
                {
                    XDocument doc;
                    using (Stream s = sheetEntry.Open())
                    {
                        doc = XDocument.Load(s);
                    }
                    XDocument rels = LoadRels(zip, sheetEntry.FullName);
                    XNamespace ns = doc.Root.Name.Namespace;
                    Dictionary<string, (int Row, int Col, string ProgId, byte[] Bytes)> byRelId = new(StringComparer.Ordinal);
                    foreach (XElement oleNode in doc.Descendants(ns + "oleObject"))
                    {
                        string rId = RelationshipId(oleNode);
                        if (string.IsNullOrEmpty(rId))
                        {
                            continue;
                        }
                        string progId = (string)oleNode.Attribute("progId");
                        XElement from = oleNode.Descendants().FirstOrDefault(n => n.Name.LocalName == "from");
                        int row = (int?)from?.Elements().FirstOrDefault(n => n.Name.LocalName == "row") ?? -1;
                        int col = (int?)from?.Elements().FirstOrDefault(n => n.Name.LocalName == "col") ?? -1;
                        if (byRelId.TryGetValue(rId, out (int Row, int Col, string ProgId, byte[] Bytes) exist))
                        {
                            // 同一对象在 mc:AlternateContent 的 Choice/Fallback 里会重复出现：保留带锚点的那份
                            if (exist.Row >= 0 || row < 0)
                            {
                                continue;
                            }
                        }
                        byte[] bytes = ResolveEmbeddingBytes(zip, rels, sheetEntry.FullName, rId);
                        if (bytes == null || bytes.Length == 0)
                        {
                            continue;
                        }
                        byRelId[rId] = (row, col, progId, bytes);
                    }
                    foreach (KeyValuePair<string, (int Row, int Col, string ProgId, byte[] Bytes)> kv in byRelId)
                    {
                        result.Add((kv.Value.Row + 1, kv.Value.Col + 1, kv.Key, kv.Value.ProgId, kv.Value.Bytes));
                    }
                }

                // 2) 新式：drawings 内的 oleObject
                foreach (ZipArchiveEntry drawingEntry in zip.Entries
                    .Where(e => e.FullName.StartsWith("xl/drawings/") && e.FullName.EndsWith(".xml")))
                {
                    XDocument doc;
                    using (Stream s = drawingEntry.Open())
                    {
                        doc = XDocument.Load(s);
                    }
                    XNamespace xdr = doc.Root.Name.Namespace;
                    XDocument rels = LoadRels(zip, drawingEntry.FullName);

                    IEnumerable<XElement> anchors = doc.Descendants(xdr + "twoCellAnchor")
                        .Concat(doc.Descendants(xdr + "oneCellAnchor"));
                    foreach (XElement anchor in anchors)
                    {
                        XElement oleNode = anchor.Descendants(xdr + "oleObject").FirstOrDefault();
                        if (oleNode == null)
                        {
                            continue;
                        }
                        int row = (int?)anchor.Element(xdr + "from")?.Element(xdr + "row") ?? -1;
                        int col = (int?)anchor.Element(xdr + "from")?.Element(xdr + "col") ?? -1;
                        string name = (string)oleNode.Attribute("name");
                        string progId = (string)oleNode.Attribute("progId");
                        string rId = RelationshipId(oleNode);
                        byte[] bytes = ResolveEmbeddingBytes(zip, rels, drawingEntry.FullName, rId);
                        if (bytes != null && bytes.Length > 0)
                        {
                            result.Add((row + 1, col + 1, name ?? rId, progId, bytes));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"解析 OLE 对象失败: {ex.Message}");
            }
            return result;
        }

        /// <summary>
        /// 读取某个部件对应的 rels（如 xl/worksheets/sheet1.xml → xl/worksheets/_rels/sheet1.xml.rels）
        /// </summary>
        private static XDocument LoadRels(ZipArchive zip, string partPath)
        {
            int slash = partPath.LastIndexOf('/');
            string relsPath = partPath.Substring(0, slash) + "/_rels/" + partPath.Substring(slash + 1) + ".rels";
            ZipArchiveEntry entry = zip.GetEntry(relsPath);
            if (entry == null)
            {
                return null;
            }
            using Stream s = entry.Open();
            return XDocument.Load(s);
        }

        /// <summary>
        /// 取元素上的关系 Id（r:id，命名空间不固定）
        /// </summary>
        private static string RelationshipId(XElement node)
            => node.Attributes().FirstOrDefault(a => a.Name.LocalName == "id" && a.Name.NamespaceName.Contains("relationships"))?.Value;

        /// <summary>
        /// 按关系 Id 从包内取 embeddings 的二进制内容
        /// </summary>
        private byte[] ResolveEmbeddingBytes(ZipArchive zip, XDocument rels, string sourcePart, string rId)
        {
            if (rels == null || string.IsNullOrEmpty(rId))
            {
                return null;
            }
            XElement rel = rels.Descendants().FirstOrDefault(n => (string)n.Attribute("Id") == rId);
            string target = (string)rel?.Attribute("Target");
            if (string.IsNullOrEmpty(target))
            {
                return null;
            }
            string binPath = ResolveRelativePartPath(sourcePart, target);
            ZipArchiveEntry binEntry = zip.GetEntry(binPath);
            if (binEntry == null)
            {
                return null;
            }
            using Stream s = binEntry.Open();
            using MemoryStream ms = new();
            s.CopyTo(ms);
            return ms.ToArray();
        }

        /// <summary>
        /// 把关系里的相对 Target 解析成包内绝对路径（如 xl/worksheets/sheet1.xml + ../embeddings/a.xlsx → xl/embeddings/a.xlsx）
        /// </summary>
        private static string ResolveRelativePartPath(string sourcePart, string target)
        {
            int slash = sourcePart.LastIndexOf('/');
            string combined = (slash >= 0 ? sourcePart.Substring(0, slash) : "") + "/" + target;
            List<string> parts = [];
            foreach (string segment in combined.Split('/'))
            {
                if (segment == "..")
                {
                    if (parts.Count > 0)
                    {
                        parts.RemoveAt(parts.Count - 1);
                    }
                }
                else if (segment != "." && segment.Length > 0)
                {
                    parts.Add(segment);
                }
            }
            return string.Join("/", parts);
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
            // 1. 中文格式：月日（年份推断）；容忍数字格式里残留的转义引号，如 1"月"9"日"
            Match m = Regex.Match(text, "([0-9]{1,2})\\s*\"?\\s*月\\s*\"?\\s*([0-9]{1,2})\\s*\"?\\s*日");
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

using NLog;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using ORT一键报告.Models;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using static ORT一键报告.Utils.Report;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 报告模板里的一项测试（含排期结果）
    /// </summary>
    public class ReportTemplateItem
    {
        /// <summary>测试项目名</summary>
        public string TestItemName { get; set; }

        /// <summary>分类（RELIABILITY TEST / EMC …）</summary>
        public string Category { get; set; }

        /// <summary>抽样计划</summary>
        public string SamplingPlan { get; set; }

        /// <summary>测试条件</summary>
        public string TestCondition { get; set; }

        /// <summary>通过判定</summary>
        public string PassCriterion { get; set; }

        /// <summary>备注</summary>
        public string Remark { get; set; }

        /// <summary>试验周期（小时，取值来自测试项目字典/计划模板）</summary>
        public string PeriodHours { get; set; }

        /// <summary>排期开始日期</summary>
        public DateTime? Start { get; set; }

        /// <summary>排期结束日期</summary>
        public DateTime? End { get; set; }

        /// <summary>周期天数（向上取整，至少 1 天）</summary>
        public int Days
        {
            get
            {
                if (Start.HasValue && End.HasValue && End.Value.Date >= Start.Value.Date)
                {
                    return (int)(End.Value.Date - Start.Value.Date).TotalDays + 1;
                }
                return PeriodDays;
            }
        }

        /// <summary>由周期小时数折算的天数</summary>
        public int PeriodDays
        {
            get
            {
                if (!double.TryParse(PeriodHours, out double hours) || hours <= 0)
                {
                    return 1;
                }
                return Math.Max(1, (int)Math.Ceiling(hours / 24.0));
            }
        }
    }

    /// <summary>
    /// 生成报告模板的输入
    /// </summary>
    public class ReportTemplateRequest
    {
        /// <summary>机种名称</summary>
        public string ModelName { get; set; }

        /// <summary>客户别（Cover 用）</summary>
        public string Customer { get; set; }

        /// <summary>版本（Cover 用）</summary>
        public string Revision { get; set; }

        /// <summary>产品阶段（Cover 用，MP/NPI 等）</summary>
        public string Stage { get; set; } = PlanStage.MP;

        /// <summary>RT 工作编号（文件夹名与报告页脚用）</summary>
        public string JobNo { get; set; }

        /// <summary>测试周期（WK2623 或 2623）</summary>
        public string Period { get; set; }

        /// <summary>序列号（每个单体一行）</summary>
        public List<string> SerialNumbers { get; set; } = [];

        /// <summary>工令</summary>
        public string WorkOrder { get; set; }

        /// <summary>单体领用/测试起始日期</summary>
        public DateTime StartDate { get; set; } = DateTime.Today;

        /// <summary>测试项（按顺序，含排期）</summary>
        public List<ReportTemplateItem> Items { get; set; } = [];

        /// <summary>输出根目录（一般是设置里的报告路径）</summary>
        public string OutputRoot { get; set; }

        /// <summary>创建人（写入 ORT Plan 的 Created By）</summary>
        public string CreatedBy { get; set; }

        /// <summary>计划备注（ORT Plan 表末 Note 一行的说明）</summary>
        public string PlanNote { get; set; }
    }

    /// <summary>
    /// 生成结果
    /// </summary>
    public class ReportTemplateResult
    {
        /// <summary>是否成功</summary>
        public bool Ok { get; set; }

        /// <summary>说明/错误原因</summary>
        public string Message { get; set; }

        /// <summary>生成的报告文件夹</summary>
        public string Folder { get; set; }

        /// <summary>生成的报告概览文件</summary>
        public string OverviewFile { get; set; }
    }

    /// <summary>
    /// 报告模板服务：以 Templates 下的报告模板为骨架，按机种的测试计划直接生成一份新的报告模板：
    /// 填 Cover 基本信息、生成 ORT Plan 表、填 Waterfall 的序列号/工令/起始日期与测试安排、按计划重建 TestStatus。
    /// 工作簿顺序固定为 Cover / ORT Plan / Waterfall / TestStatus。
    /// </summary>
    public class ReportTemplateService
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;

        /// <summary>EMU/像素（Excel 锚点单位：914400 EMU = 1 英寸 = 96 像素）</summary>
        private const int EmuPerPixel = 9525;

        /// <summary>测试安排区外框颜色（黑）</summary>
        private const short IndexedBlackBorder = 8;

        /// <summary>测试安排区内线颜色（50% 灰）</summary>
        private const short IndexedGreyBorder = 23;

        /// <summary>测试项配图在 ORT Plan 表里的最大尺寸（像素），超出按比例缩小</summary>
        private const int PictureMaxWidth = 200;
        private const int PictureMaxHeight = 110;

        /// <summary>ORT Plan 里测试项配图所在的列（H 列，表格右侧的照片列，不压文字）</summary>
        private const int PictureColumn = 8;

        /// <summary>每个测试项最多放几张配图</summary>
        private const int PictureMaxPerItem = 2;

        /// <summary>ORT Plan 大标题颜色 RGB(0,0,204)</summary>
        private static readonly byte[] OrtPlanTitleRgb = [0x00, 0x00, 0xCC];

        /// <summary>ORT Plan 测试种类行底纹 RGB(197,217,241)</summary>
        private static readonly byte[] OrtPlanCategoryRgb = [197, 217, 241];

        /// <summary>TEST CONDITOIN / PASS CRITERION 是否把冒号前的词组加粗</summary>
        private const bool BoldBeforeColon = true;

        public ReportTemplateService() { }

        public ReportTemplateService(DatabaseService db)
        {
            _db = db;
        }

        /// <summary>模板目录名（程序目录 Templates 下）</summary>
        public const string TemplateFolderName = "# ORT Test Report (WK#)_#";

        /// <summary>模板目录完整路径</summary>
        public static string TemplateDir
            => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", TemplateFolderName);

        /* ###############################  排期  ################################ */

        /// <summary>
        /// 按"每项测试的试验周期、开始日期落在工作日"给测试项排期。
        /// 每项测试占连续的若干天（周期小时数 / 24 向上取整），下一项从上一项结束后的第一个工作日开始。
        /// </summary>
        public static void Schedule(List<ReportTemplateItem> items, DateTime startDate)
        {
            if (items == null || items.Count == 0)
            {
                return;
            }
            DateTime cursor = startDate.Date;
            foreach (ReportTemplateItem item in items)
            {
                int days = item.PeriodDays;
                DateTime start = NextWorkingDay(cursor);
                DateTime end = start.AddDays(days - 1);
                item.Start = start;
                item.End = end;
                cursor = end.AddDays(1);
            }
        }

        /// <summary>日期所在（或之后的）第一个工作日（跳过周六周日）</summary>
        public static DateTime NextWorkingDay(DateTime date)
        {
            DateTime value = date.Date;
            while (value.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday)
            {
                value = value.AddDays(1);
            }
            return value;
        }

        /* ###############################  生成  ################################ */

        /// <summary>
        /// 生成报告模板：复制模板文件夹 → 写入 Cover/ORT Plan/Waterfall/TestStatus
        /// </summary>
        public ReportTemplateResult Generate(ReportTemplateRequest request)
        {
            ReportTemplateResult result = new();
            if (request == null || string.IsNullOrWhiteSpace(request.ModelName))
            {
                result.Message = "机种名称不能为空";
                return result;
            }
            if (string.IsNullOrWhiteSpace(request.OutputRoot) || !Directory.Exists(request.OutputRoot))
            {
                result.Message = "输出目录不存在";
                return result;
            }
            if (!Directory.Exists(TemplateDir))
            {
                result.Message = $"找不到报告模板：{TemplateDir}";
                return result;
            }
            try
            {
                // 排期：调用方一般已经排好了，这里再排一次（幂等），避免漏排时 Waterfall 是空的
                Schedule(request.Items, request.StartDate);
                string week = ExtractWeek(request.Period);
                string folderName = BuildFolderName(request.ModelName, week, request.JobNo);
                string targetFolder = Path.Combine(request.OutputRoot, folderName);
                CopyDirectory(TemplateDir, targetFolder);

                string templateWorkbook = Directory.GetFiles(targetFolder, "*.xls*")
                    .Where(f => !Path.GetFileName(f).StartsWith("~$", StringComparison.Ordinal))
                    .OrderByDescending(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                    .FirstOrDefault();
                if (templateWorkbook == null)
                {
                    result.Message = "模板文件夹里没有找到报告概览 Excel";
                    return result;
                }
                string overviewName = $"{request.ModelName} ORT Test Report (WK{week}).xlsx";
                string overviewPath = Path.Combine(targetFolder, overviewName);
                if (!string.Equals(templateWorkbook, overviewPath, StringComparison.OrdinalIgnoreCase))
                {
                    if (File.Exists(overviewPath))
                    {
                        File.Delete(overviewPath);
                    }
                    File.Move(templateWorkbook, overviewPath);
                }

                XSSFWorkbook wb = ExcelNpoi.OpenRead(overviewPath);
                try
                {
                    // 模板是从历史报告另存来的，里面还挂着共享公式组；
                    // 先规范化（写成普通公式），否则删/插行后共享公式失效，Excel 打开会要求修复
                    ExcelNpoi.NormalizeFormulas(wb);
                    WriteCover(wb, request);
                    WriteOrtPlan(wb, request);
                    WriteWaterfall(wb, request);
                    WriteTestStatus(wb, request, targetFolder);
                    EnsureSheetOrder(wb);
                    wb.SetForceFormulaRecalculation(true);
                    ExcelNpoi.Save(wb, overviewPath);
                }
                finally
                {
                    wb.Close();
                }

                // 保存后清掉外部工作簿引用/失效命名区域/calcChain，避免 Excel 打开时提示"检测到错误"
                XlsxCleaner.Clean(overviewPath);

                result.Ok = true;
                result.Folder = targetFolder;
                result.OverviewFile = overviewPath;
                result.Message = overviewPath;
                _logger.Info($"报告模板已生成：{overviewPath}");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "生成报告模板失败");
                result.Message = ex.Message;
            }
            return result;
        }

        /// <summary>报告文件夹名：{机种} ORT Test Report (WK{周号})_{RT工号}</summary>
        public static string BuildFolderName(string modelName, string week, string jobNo)
        {
            string name = $"{modelName} ORT Test Report (WK{week})";
            return string.IsNullOrWhiteSpace(jobNo) ? name : $"{name}_{jobNo.Trim()}";
        }

        /// <summary>从 WK2623 / 2623 里取 4 位周号；取不到用当前年份+周序号</summary>
        public static string ExtractWeek(string period)
        {
            Match match = Regex.Match(period ?? "", @"(\d{4})");
            if (match.Success)
            {
                return match.Groups[1].Value;
            }
            return DateTime.Today.Year.ToString("0000").Substring(2) + GetIsoWeek(DateTime.Today).ToString("00");
        }

        /// <summary>ISO 周号（.NET Framework 没有 ISOWeek）</summary>
        public static int GetIsoWeek(DateTime date)
        {
            DayOfWeek day = System.Globalization.CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(date);
            if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday)
            {
                date = date.AddDays(3);
            }
            return System.Globalization.CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(
                date, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        }

        /// <summary>
        /// 递归复制目录（跳过 Excel 临时文件与遗留的 setup_info.json）
        /// </summary>
        private static void CopyDirectory(string source, string target)
        {
            Directory.CreateDirectory(target);
            foreach (string file in Directory.GetFiles(source))
            {
                string name = Path.GetFileName(file);
                if (name.StartsWith("~$", StringComparison.Ordinal) || IsLegacyFile(name))
                {
                    continue;
                }
                File.Copy(file, Path.Combine(target, name), true);
            }
            foreach (string dir in Directory.GetDirectories(source))
            {
                CopyDirectory(dir, Path.Combine(target, Path.GetFileName(dir)));
            }
        }

        /// <summary>
        /// 早期模板里遗留、生成报告时不需要的文件（setup_info.json 是旧版设置文件，已废弃）
        /// </summary>
        private static bool IsLegacyFile(string fileName)
            => fileName.Equals("setup_info.json", StringComparison.OrdinalIgnoreCase);

        /* ###############################  Cover  ################################ */

        /// <summary>
        /// Cover：Customer / Model Name / Revision / Product Stage 写在 J12/J14/J16/J18
        /// （与 setup_info.json 及已完成报告一致），日期一类公式保持原样自动计算。
        /// </summary>
        private static void WriteCover(IWorkbook wb, ReportTemplateRequest request)
        {
            ISheet sheet = FindSheet(wb, "cover") ?? ExcelNpoi.SheetAt(wb, 0);
            if (sheet == null)
            {
                return;
            }
            ExcelNpoi.SetCell(sheet, 12, 10, request.Customer ?? "");
            ExcelNpoi.SetCell(sheet, 14, 10, request.ModelName ?? "");
            ExcelNpoi.SetCell(sheet, 16, 10, request.Revision ?? "");
            ExcelNpoi.SetCell(sheet, 18, 10, string.IsNullOrWhiteSpace(request.Stage) ? PlanStage.MP : request.Stage);
        }

        /* ###############################  ORT Plan  ################################ */

        /// <summary>
        /// 生成 ORT Plan 表：标题 + 建立信息 + 表头 + 分类行/测试项行 + 表末 Note。
        /// 版式与原报告一致（B=NO. C=TEST ITEMS D=SAMPLING PLAN E=TEST CONDITOIN F=PASS CRITERION G=COMMENT）。
        /// </summary>
        private void WriteOrtPlan(IWorkbook wb, ReportTemplateRequest request)
        {
            ISheet sheet = FindSheet(wb, "ort plan") ?? wb.CreateSheet("ORT Plan");
            // 清空重建（模板里没有这张表，或已有内容时都按计划重写）
            for (int r = sheet.LastRowNum; r >= 0; r--)
            {
                IRow existingRow = sheet.GetRow(r);
                if (existingRow != null)
                {
                    sheet.RemoveRow(existingRow);
                }
            }
            for (int i = sheet.NumMergedRegions - 1; i >= 0; i--)
            {
                sheet.RemoveMergedRegion(i);
            }

            ExcelNpoi.SetCell(sheet, 1, 4, "Ongoing Reliability Test Plan");
            // 大标题：加粗、24 号、颜色 RGB(0,0,204)（与 Cover/Waterfall/TestStatus 的标题同色）
            ExcelNpoi.ApplyStyle(sheet, 1, 4, 1, 4, new ExcelNpoi.CellStyleSpec
            {
                Bold = true,
                FontSize = 24,
                FontRgb = OrtPlanTitleRgb,
                Vertical = VerticalAlignment.Center
            });
            ExcelNpoi.SetCell(sheet, 3, 2, string.IsNullOrWhiteSpace(request.CreatedBy) ? "" : $"Created By: {request.CreatedBy}");
            ExcelNpoi.SetCell(sheet, 3, 4, "Model: ");
            ExcelNpoi.SetCell(sheet, 3, 5, request.ModelName ?? "");

            string[] headers = ["NO.", "TEST\nITEMS", "SAMPLING PLAN", "TEST CONDITOIN", "PASS CRITERION", "COMMENT"];
            for (int i = 0; i < headers.Length; i++)
            {
                ExcelNpoi.SetCell(sheet, 4, 2 + i, headers[i]);
            }
            // 表头：黑底白字、加粗、居中
            ExcelNpoi.ApplyStyle(sheet, 4, 2, 4, 7, new ExcelNpoi.CellStyleSpec
            {
                Bold = true,
                Border = BorderStyle.Thin,
                Horizontal = HorizontalAlignment.Center,
                Vertical = VerticalAlignment.Center,
                WrapText = true,
                FontRgb = [0xFF, 0xFF, 0xFF],
                FillRgb = [0x00, 0x00, 0x00]
            });
            // 表头与第一条分类之间的空行压到 5 像素高
            ExcelNpoi.SetRowHeight(sheet, 5, 5 * 72.0 / 96.0);

            int row = 6;
            int no = 0;
            int categoryNo = 0;
            List<(ReportTemplateItem Item, int Row)> itemRows = [];
            // 按"测试种类"分组（同一类的测试项排在一起，历史报告就是这个样子）
            foreach (KeyValuePair<string, List<ReportTemplateItem>> group in GroupByCategory(request.Items))
            {
                categoryNo++;
                no = 0;
                ExcelNpoi.SetCell(sheet, row, 2, categoryNo);
                ExcelNpoi.SetCell(sheet, row, 3, group.Key);
                // 分类行：淡蓝底（RGB 197,217,241）、加粗、居中
                ExcelNpoi.ApplyStyle(sheet, row, 2, row, 7, new ExcelNpoi.CellStyleSpec
                {
                    Bold = true,
                    Border = BorderStyle.Thin,
                    Horizontal = HorizontalAlignment.Center,
                    Vertical = VerticalAlignment.Center,
                    FillRgb = OrtPlanCategoryRgb
                });
                ExcelNpoi.Merge(sheet, row, 3, row, 7);
                row++;

                foreach (ReportTemplateItem item in group.Value)
                {
                    no++;
                    ExcelNpoi.SetCell(sheet, row, 2, no);
                    ExcelNpoi.SetCell(sheet, row, 3, item.TestItemName ?? "");
                    ExcelNpoi.SetCell(sheet, row, 4, item.SamplingPlan ?? "");
                    ExcelNpoi.SetRichText(sheet, row, 5, item.TestCondition, BoldBeforeColon);
                    ExcelNpoi.SetRichText(sheet, row, 6, item.PassCriterion, BoldBeforeColon);
                    ExcelNpoi.SetCell(sheet, row, 7, item.Remark ?? "");
                    // 序号与测试名称居中，内容列左对齐并自动换行
                    ExcelNpoi.ApplyStyle(sheet, row, 2, row, 3, new ExcelNpoi.CellStyleSpec
                    {
                        Border = BorderStyle.Thin,
                        Horizontal = HorizontalAlignment.Center,
                        Vertical = VerticalAlignment.Center,
                        WrapText = true
                    });
                    ExcelNpoi.ApplyStyle(sheet, row, 4, row, 7, new ExcelNpoi.CellStyleSpec
                    {
                        Border = BorderStyle.Thin,
                        Vertical = VerticalAlignment.Top,
                        WrapText = true
                    });
                    itemRows.Add((item, row));
                    row++;
                }
            }

            // 表末 Note（计划的抽样原则说明；计划里没有就用通用文案）
            row++;
            ExcelNpoi.SetCell(sheet, row, 3, string.IsNullOrWhiteSpace(request.PlanNote)
                ? "Note.\nSampling principle: samples plan should cover the different product lines and shifts."
                : $"Note.\n{request.PlanNote}");
            ExcelNpoi.ApplyStyle(sheet, row, 3, row, 7, new ExcelNpoi.CellStyleSpec { WrapText = true, Vertical = VerticalAlignment.Top });
            ExcelNpoi.Merge(sheet, row, 3, row, 7);

            ExcelNpoi.SetColumnWidth(sheet, 1, 1.0);
            ExcelNpoi.SetColumnWidth(sheet, 2, 4.5);
            ExcelNpoi.SetColumnWidth(sheet, 3, 19.5);
            ExcelNpoi.SetColumnWidth(sheet, 4, 16);
            ExcelNpoi.SetColumnWidth(sheet, 5, 31.4);
            ExcelNpoi.SetColumnWidth(sheet, 6, 37.9);
            ExcelNpoi.SetColumnWidth(sheet, 7, 12.4);
            // 照片列（H）：测试项配图放在这里，宽度按配图最大宽度留够
            ExcelNpoi.SetColumnWidth(sheet, PictureColumn, PictureMaxWidth / 7.0 + 2);

            WriteOrtPlanLogo(wb, sheet);
            WriteOrtPlanPictures(wb, sheet, itemRows);
        }

        /// <summary>
        /// ORT Plan 标题左边的公司 logo：其他三张表都有，模板里没有这张表，
        /// 这里从同一工作簿的其他表里取（最小的那张图就是 logo），按历史报告的锚点放在标题左边。
        /// </summary>
        private static void WriteOrtPlanLogo(IWorkbook wb, ISheet sheet)
        {
            if (sheet == null || sheet.CreateDrawingPatriarch() is XSSFDrawing drawing && drawing.GetShapes().Count > 0)
            {
                return; // 已经有图（例如模板本身就带 logo）就不重复放
            }
            byte[] logo = FindLogoBytes(wb);
            if (logo == null)
            {
                return;
            }
            ExcelNpoi.AddPictureAnchored(wb, sheet, logo, ExcelNpoi.DetectPictureType(logo),
                col1: 2, row1: 1, col2: 3, row2: 2, dx1: 0, dy1: 19050, dx2: 1104900, dy2: 190500);
        }

        /// <summary>找公司 logo 的图片字节：整册里体积最小的那张图（其他三张表标题左边都是它）</summary>
        private static byte[] FindLogoBytes(IWorkbook wb)
        {
            byte[] best = null;
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                foreach (ExcelNpoi.SheetPicture picture in ExcelNpoi.PictureDetails(wb.GetSheetAt(i)))
                {
                    byte[] bytes = picture.Bytes;
                    if (bytes == null || bytes.Length == 0 || bytes.Length > 20 * 1024)
                    {
                        continue;
                    }
                    if (best == null || bytes.Length < best.Length)
                    {
                        best = bytes;
                    }
                }
            }
            return best;
        }

        /// <summary>
        /// 让每个测试项的合并跨度都放得下它的名字：跨度总宽不够时，把差额平摊到跨度内的各列
        /// （历史报告的日期列宽度就是按内容拉开的，否则文字会溢出、和相邻测试项重叠）
        /// </summary>
        private static void EnsureItemColumnWidths(ISheet sheet, List<(int From, int To, string Text)> spans, int firstCol, int lastCol)
        {
            if (sheet == null || spans.Count == 0 || lastCol < firstCol)
            {
                return;
            }
            double[] widths = new double[lastCol + 1];
            for (int c = firstCol; c <= lastCol; c++)
            {
                widths[c] = ExcelNpoi.ColumnWidthInChars(sheet, c);
            }
            foreach ((int from, int to, string text) in spans)
            {
                int a = Math.Max(from, firstCol);
                int b = Math.Min(to, lastCol);
                if (b < a || string.IsNullOrEmpty(text))
                {
                    continue;
                }
                double total = 0;
                for (int c = a; c <= b; c++)
                {
                    total += widths[c];
                }
                double needed = MeasureTextWidth(text) + 2.0; // 两侧各留一点边距
                if (needed <= total)
                {
                    continue;
                }
                double extra = (needed - total) / (b - a + 1);
                for (int c = a; c <= b; c++)
                {
                    widths[c] += extra;
                }
            }
            for (int c = firstCol; c <= lastCol; c++)
            {
                ExcelNpoi.SetColumnWidth(sheet, c, widths[c]);
            }
        }

        /// <summary>
        /// 估算文字占用的字符宽（列宽单位）：全角字符按 2、其余按 1
        /// </summary>
        private static double MeasureTextWidth(string text)
        {
            double width = 0;
            foreach (char ch in text ?? "")
            {
                width += ch > 0x2E80 ? 2.0 : 1.0;
            }
            return width;
        }

        /// <summary>
        /// 测试项配图：计划索引时从历史报告的 ORT Plan 表里抽出来并按测试项归好了类，
        /// 这里把同名测试项的图片放在该测试项所在行的**照片列（H 列，表格右侧）**，
        /// 不再压到 TEST CONDITOIN / PASS CRITERION 的文字上；并按图片高度把行高撑开，避免相邻项目互相遮挡。
        /// </summary>
        private void WriteOrtPlanPictures(IWorkbook wb, ISheet sheet, List<(ReportTemplateItem Item, int Row)> itemRows)
        {
            if (_db == null || itemRows.Count == 0)
            {
                return;
            }
            try
            {
                foreach ((ReportTemplateItem item, int row) in itemRows)
                {
                    string key = PlanIndexService.NameKey(item.TestItemName);
                    if (string.IsNullOrEmpty(key))
                    {
                        continue;
                    }
                    List<PlanItemImage> images = _db.FreeSql.Select<PlanItemImage>()
                        .Where(i => i.NameKey == key)
                        .OrderBy(i => i.OrderNo)
                        .ToList();
                    if (images.Count == 0)
                    {
                        continue;
                    }
                    int offsetY = 0;
                    int placed = 0;
                    foreach (PlanItemImage image in images)
                    {
                        if (placed >= PictureMaxPerItem)
                        {
                            break;
                        }
                        string path = Path.Combine(_db.PlanImagesDir, image.FileName ?? "");
                        if (!File.Exists(path))
                        {
                            continue;
                        }
                        byte[] bytes = File.ReadAllBytes(path);
                        ScaleToFit(image.WidthPx, image.HeightPx, out int width, out int height);
                        ExcelNpoi.AddPictureAnchored(wb, sheet, bytes, ExcelNpoi.DetectPictureType(bytes),
                            col1: PictureColumn, row1: row, col2: PictureColumn, row2: row,
                            dx1: 0, dy1: offsetY,
                            dx2: width * EmuPerPixel, dy2: offsetY + height * EmuPerPixel);
                        offsetY += (height + 4) * EmuPerPixel;
                        placed++;
                    }
                    if (placed > 0)
                    {
                        // 行高取"文字估算高度"与"图片总高"的较大值，保证图片不被下一行压住、文字也不被裁掉
                        double imagesHeight = offsetY / (double)EmuPerPixel * 72.0 / 96.0;
                        double textHeight = EstimateOrtPlanRowHeight(item);
                        ExcelNpoi.SetRowHeight(sheet, row, Math.Max(imagesHeight + 4, textHeight));
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"写入 ORT Plan 配图失败：{ex.Message}");
            }
        }

        /// <summary>按最大尺寸等比缩放（尺寸读不出来时给默认值）</summary>
        private static void ScaleToFit(int widthPx, int heightPx, out int width, out int height)
        {
            width = widthPx > 0 ? widthPx : PictureMaxWidth;
            height = heightPx > 0 ? heightPx : PictureMaxHeight;
            double scale = Math.Min(1.0, Math.Min((double)PictureMaxWidth / width, (double)PictureMaxHeight / height));
            width = Math.Max(1, (int)Math.Round(width * scale));
            height = Math.Max(1, (int)Math.Round(height * scale));
        }

        /// <summary>
        /// 估算 ORT Plan 某个测试项行的文字高度（磅）：
        /// 按各列的字符宽度折行数（D=16 / E=31.4 / F=37.9 字符），取最多的那一列
        /// </summary>
        private static double EstimateOrtPlanRowHeight(ReportTemplateItem item)
        {
            double lines = 1;
            lines = Math.Max(lines, Lines(item.SamplingPlan, 16));
            lines = Math.Max(lines, Lines(item.TestCondition, 31.4));
            lines = Math.Max(lines, Lines(item.PassCriterion, 37.9));
            lines = Math.Max(lines, Lines(item.Remark, 12.4));
            return lines * 13.5 + 4;

            static double Lines(string text, double charsPerLine)
            {
                if (string.IsNullOrEmpty(text) || charsPerLine <= 0)
                {
                    return 1;
                }
                double total = 0;
                foreach (string line in text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'))
                {
                    total += Math.Max(1, Math.Ceiling(MeasureTextWidth(line) / charsPerLine));
                }
                return total;
            }
        }

        /* ###############################  Waterfall  ################################ */

        /// <summary>
        /// Waterfall：D 列填序列号、序列号之后每行一个工令、起始日期写在第一列（其余列由公式顺推，
        /// 格式沿用模板里日期/星期/UUT 序号三行的既有格式），右侧按排期写入每天的测试安排
        /// （同一天范围的同名测试合并单元格，单元格格式沿用模板里的测试项行）。
        /// </summary>
        private static void WriteWaterfall(IWorkbook wb, ReportTemplateRequest request)
        {
            ISheet sheet = FindSheet(wb, "waterfall") ?? ExcelNpoi.SheetAt(wb, 2);
            if (sheet == null)
            {
                return;
            }
            const int weekRow = 3;   // 星期（=WEEKDAY(日期,2)，周末由模板的条件格式涂灰）
            const int dateRow = 4;   // 日期（第一列填起始日期，其余 =前一列+1）
            const int uutRow = 5;    // UUT S/N 序号（第一列 1，其余 =前一列+1）
            const int firstSnRow = 7;
            const int firstDateCol = 5; // E 列
            const int templateSnRows = 3;

            // 模板里这三行与测试项行各自的格式先取下来（清空后样式就没了）
            ICellStyle weekStyle = ExcelNpoi.ExistingCell(sheet, weekRow, firstDateCol)?.CellStyle;
            ICellStyle dateStyle = ExcelNpoi.ExistingCell(sheet, dateRow, firstDateCol)?.CellStyle;
            ICellStyle uutStyle = ExcelNpoi.ExistingCell(sheet, uutRow, firstDateCol)?.CellStyle;
            ICellStyle itemStyle = ExcelNpoi.ExistingCell(sheet, firstSnRow, firstDateCol)?.CellStyle;
            ICellStyle lastItemStyle = ExcelNpoi.ExistingCell(sheet, firstSnRow + templateSnRows - 1, firstDateCol)?.CellStyle;
            int templateLastCol = Math.Max(ExcelNpoi.LastColumn(sheet), firstDateCol);

            List<string> sns = (request.SerialNumbers ?? []).Where(s => !string.IsNullOrWhiteSpace(s)).Select(s => s.Trim()).ToList();
            int snCount = Math.Max(sns.Count, 1);
            List<string> workOrders = (request.WorkOrder ?? "")
                .Split(['\n', '\r', '\t', ' ', ',', '，', ';', '；'], StringSplitOptions.RemoveEmptyEntries)
                .Select(w => w.Trim())
                .Where(w => w.Length > 0)
                .ToList();

            // 1. 序列号行数按实际调整（模板里预留了 3 行）
            if (snCount < templateSnRows)
            {
                ExcelNpoi.DeleteRows(sheet, firstSnRow + snCount, templateSnRows - snCount);
            }
            else if (snCount > templateSnRows)
            {
                int extra = snCount - templateSnRows;
                for (int i = 0; i < extra; i++)
                {
                    int insertAt = firstSnRow + templateSnRows + i;
                    ExcelNpoi.InsertRows(sheet, insertAt, 1);
                    ExcelNpoi.CopyRowStyle(sheet, firstSnRow + templateSnRows - 1, insertAt, templateLastCol);
                }
            }

            // 2. 清空并重建"测试安排"区域（先去掉该区域的合并，避免残留）；
            //    表头块（No. / UUT S/N / Config / S/N，行 3~5 的 B~D 列）与它的合并要保留，
            //    所以合并只从第 6 行开始清、这三行也只从 E 列开始清空
            int lastRow = ExcelNpoi.LastRow(sheet);
            int clearLastRow = firstSnRow + snCount + Math.Max(workOrders.Count, 1); // 含工令行
            int lastCol = Math.Max(templateLastCol, firstDateCol + snCount * 20);
            RemoveMergesInRows(sheet, firstSnRow - 1, Math.Max(clearLastRow, lastRow));
            for (int r = weekRow; r <= clearLastRow; r++)
            {
                bool headerRow = r <= uutRow;
                if (headerRow)
                {
                    // 表头三行只清日程区（E 列起），"UUT S/N / Config / S/N" 这些标题留着
                    if (lastCol >= firstDateCol)
                    {
                        ExcelNpoi.ClearCells(sheet, r, firstDateCol, lastCol);
                    }
                    continue;
                }
                ExcelNpoi.ClearCells(sheet, r, 3, lastCol);
                if (r >= firstSnRow && r < firstSnRow + snCount)
                {
                    ExcelNpoi.SetCell(sheet, r, 2, r - firstSnRow + 1);
                }
            }

            // 3. 起始日期 + 序列号 + 工令（多个工令各占一行）
            //    起始日期对齐到工作日：第一项测试必然落在 E 列，概览读取（Waterfall 的 S/N 判定）也才稳
            DateTime start = NextWorkingDay(request.StartDate);
            for (int i = 0; i < sns.Count; i++)
            {
                ExcelNpoi.SetCell(sheet, firstSnRow + i, 4, sns[i]);
            }
            for (int i = 0; i < workOrders.Count; i++)
            {
                ExcelNpoi.SetCell(sheet, firstSnRow + snCount + i, 4, workOrders[i]);
            }

            // 4. 日期/星期/UUT 序号：第一列填起始日期，其余列用公式顺推（格式沿用模板列）
            int totalDays = 0;
            foreach (ReportTemplateItem item in request.Items ?? [])
            {
                if (item.Start.HasValue && item.End.HasValue)
                {
                    totalDays = Math.Max(totalDays, (int)(item.End.Value.Date - start).TotalDays + 1);
                }
            }
            totalDays = Math.Max(totalDays, 1);
            int lastDateCol = firstDateCol + totalDays - 1;
            for (int c = firstDateCol; c <= lastDateCol; c++)
            {
                ICell dateCell = ExcelNpoi.Cell(sheet, dateRow, c);
                if (c == firstDateCol)
                {
                    dateCell.SetCellValue(start);
                }
                else
                {
                    dateCell.CellFormula = $"{ExcelNpoi.AddressOf(dateRow, c - 1)}+1";
                }
                if (dateStyle != null)
                {
                    dateCell.CellStyle = dateStyle;
                }

                ICell weekCell = ExcelNpoi.Cell(sheet, weekRow, c);
                weekCell.CellFormula = $"WEEKDAY({ExcelNpoi.AddressOf(dateRow, c)},2)";
                if (weekStyle != null)
                {
                    weekCell.CellStyle = weekStyle;
                }

                ICell uutCell = ExcelNpoi.Cell(sheet, uutRow, c);
                if (c == firstDateCol)
                {
                    uutCell.SetCellValue(1d);
                }
                else
                {
                    uutCell.CellFormula = $"{ExcelNpoi.AddressOf(uutRow, c - 1)}+1";
                }
                if (uutStyle != null)
                {
                    uutCell.CellStyle = uutStyle;
                }
            }
            // 排期结束之后的列清空（模板里留着上一条报告的日期，别留在新表上）
            for (int c = lastDateCol + 1; c <= templateLastCol; c++)
            {
                for (int r = weekRow; r <= uutRow; r++)
                {
                    ExcelNpoi.ClearCells(sheet, r, c, c);
                }
            }
            for (int c = templateLastCol + 1; c <= lastDateCol; c++)
            {
                ExcelNpoi.SetColumnWidth(sheet, c, 5.4);
            }

            // 5. 测试安排：每个序列号都走一遍计划里的测试（同一天范围合并），单元格格式沿用模板的测试项行
            int lastSnRow = firstSnRow + snCount - 1;
            List<(int From, int To, string Text)> spans = [];
            foreach (ReportTemplateItem item in request.Items ?? [])
            {
                if (!item.Start.HasValue || !item.End.HasValue)
                {
                    continue;
                }
                int from = firstDateCol + (int)(item.Start.Value.Date - start).TotalDays;
                int to = firstDateCol + (int)(item.End.Value.Date - start).TotalDays;
                if (from < firstDateCol)
                {
                    continue;
                }
                if (to > lastDateCol)
                {
                    to = lastDateCol;
                }
                spans.Add((from, to, item.TestItemName ?? ""));
                for (int i = 0; i < snCount; i++)
                {
                    int row = firstSnRow + i;
                    ExcelNpoi.SetCell(sheet, row, from, item.TestItemName ?? "");
                    ICellStyle style = row == lastSnRow && lastItemStyle != null ? lastItemStyle : itemStyle;
                    if (style != null)
                    {
                        ExcelNpoi.ApplyStyle(sheet, row, from, row, to, style);
                    }
                    else
                    {
                        ExcelNpoi.ApplyStyle(sheet, row, from, row, to, new ExcelNpoi.CellStyleSpec
                        {
                            Horizontal = HorizontalAlignment.Center,
                            Vertical = VerticalAlignment.Center,
                            WrapText = true
                        });
                    }
                    if (to > from)
                    {
                        ExcelNpoi.Merge(sheet, row, from, row, to);
                    }
                }
            }

            // 5.5 列宽按测试项文字自适应：合并跨度里放不下测试项目名时会溢出到相邻单元格造成重叠，
            //     这里按"跨度总宽 ≥ 文字宽"补足（历史报告里日期列宽度就是这样按内容拉开的）
            EnsureItemColumnWidths(sheet, spans, firstDateCol, lastDateCol);

            // 6. 排期用不到的末尾列：整列清干净（连模板残留的黑底/边框一起去掉）
            ICellStyle blankStyle = ExcelNpoi.ExistingCell(sheet, weekRow - 1, firstDateCol)?.CellStyle
                ?? ExcelNpoi.BorderStyleOf(wb, BorderStyle.None);
            for (int c = lastDateCol + 1; c <= templateLastCol; c++)
            {
                for (int r = weekRow; r <= clearLastRow; r++)
                {
                    ExcelNpoi.ClearCells(sheet, r, c, c);
                    ExcelNpoi.Cell(sheet, r, c).CellStyle = blankStyle;
                }
            }

            // 7. 测试项目安排区整体加框：B7 到"最后一天 + 最后一个序列号"，
            //    内部细灰线、最外一圈粗黑线（历史报告就是这个版式）
            ExcelNpoi.ApplyBlockBorder(wb, sheet, firstSnRow, 2, lastSnRow, lastDateCol,
                BorderStyle.Thin, BorderStyle.Medium, IndexedGreyBorder, IndexedBlackBorder);
        }

        /* ###############################  TestStatus  ################################ */

        /// <summary>
        /// TestStatus：按计划重建"分类行 + 测试项行"，并重算合计行。
        /// 分类取测试项目的"测试种类"（可靠性/环境 → ENVIRONMENT TESTS，EMC → EMC ），
        /// 每个测试项的单元格样式、超链接（指向 Report 文件夹里对应的报告）与整表底色都按历史报告的样子还原。
        /// </summary>
        private static void WriteTestStatus(IWorkbook wb, ReportTemplateRequest request, string targetFolder)
        {
            ISheet sheet = FindSheet(wb, "teststatus") ?? FindSheet(wb, "test status");
            if (sheet == null)
            {
                return;
            }
            const int totalRow = 8;
            const int firstDataRow = 9;
            // 模板里 合计行 / 分类行 / 测试项行 分别留了各列的格式，逐列取下来套用
            ICellStyle[] totalStyles = CaptureRowStyles(sheet, totalRow);
            ICellStyle[] categoryStyles = CaptureRowStyles(sheet, firstDataRow);
            ICellStyle[] itemStyles = CaptureRowStyles(sheet, firstDataRow + 1);
            int templateLast = ExcelNpoi.LastRow(sheet);
            int templateRows = Math.Max(templateLast - firstDataRow + 1, 0);

            // 用列表保留分类的先后顺序
            List<KeyValuePair<string, List<ReportTemplateItem>>> groups = [];
            foreach (ReportTemplateItem item in request.Items ?? [])
            {
                string category = TestCategories.StatusDisplay(item.Category);
                int index = groups.FindIndex(g => string.Equals(g.Key, category, StringComparison.OrdinalIgnoreCase));
                if (index < 0)
                {
                    groups.Add(new KeyValuePair<string, List<ReportTemplateItem>>(category, [item]));
                }
                else
                {
                    groups[index].Value.Add(item);
                }
            }
            int need = groups.Sum(g => 1 + g.Value.Count);
            if (templateRows > 0)
            {
                ExcelNpoi.DeleteRows(sheet, firstDataRow, templateRows);
            }
            if (need > 0)
            {
                ExcelNpoi.InsertRows(sheet, firstDataRow, need);
            }
            RemoveMergesInRows(sheet, totalRow, Math.Max(firstDataRow, ExcelNpoi.LastRow(sheet)));

            // 合计行（B8:D8 合并）
            ExcelNpoi.SetCell(sheet, totalRow, 2, "Total");
            ExcelNpoi.Merge(sheet, totalRow, 2, totalRow, 4);
            ApplyRowStyles(sheet, totalRow, totalStyles);

            int row = firstDataRow;
            int categoryNo = 0;
            List<int> categoryRows = [];
            List<(int Row, string ItemName)> itemNameRows = [];
            foreach (KeyValuePair<string, List<ReportTemplateItem>> group in groups)
            {
                categoryNo++;
                categoryRows.Add(row);
                int firstItemRow = row + 1;
                int lastItemRow = row + group.Value.Count;
                ExcelNpoi.SetCell(sheet, row, 2, categoryNo);
                ExcelNpoi.SetCell(sheet, row, 4, group.Key);
                ExcelNpoi.Merge(sheet, row, 2, row, 3);
                ApplyRowStyles(sheet, row, categoryStyles);
                SumRow(sheet, row, firstItemRow, lastItemRow);
                row++;

                int itemNo = 0;
                foreach (ReportTemplateItem item in group.Value)
                {
                    itemNo++;
                    ExcelNpoi.SetCell(sheet, row, 2, itemNo);
                    ExcelNpoi.SetCell(sheet, row, 3, $"{categoryNo}.{itemNo}");
                    ExcelNpoi.SetCell(sheet, row, 4, item.TestItemName ?? "");
                    ExcelNpoi.SetCell(sheet, row, 5, request.SerialNumbers?.Count ?? 0);
                    // 已测/失败/未测先置 0，测试过程中由使用者填写；比率用 IFERROR 兜住除零
                    ExcelNpoi.SetCell(sheet, row, 6, 0d);
                    ExcelNpoi.SetCell(sheet, row, 7, 0d);
                    ExcelNpoi.SetCell(sheet, row, 8, 0d);
                    ApplyRowStyles(sheet, row, itemStyles);
                    // 完成率/% 与合计行同构：由已录入的完成数自动计算
                    SetFormula(sheet, row, 9, $"IFERROR((F{row}+H{row})/E{row},0)");
                    SetFormula(sheet, row, 10, $"IFERROR(G{row}/F{row},0)");
                    SetFormula(sheet, row, 11, $"IFERROR((F{row}-G{row})/E{row},0)");
                    itemNameRows.Add((row, item.TestItemName ?? ""));
                    row++;
                }
            }

            // 合计行引用各分类行
            SetFormula(sheet, totalRow, 5, $"SUM({string.Join(",", categoryRows.Select(r => $"E{r}"))})");
            SetFormula(sheet, totalRow, 6, $"SUM({string.Join(",", categoryRows.Select(r => $"F{r}"))})");
            SetFormula(sheet, totalRow, 7, $"SUM({string.Join(",", categoryRows.Select(r => $"G{r}"))})");
            SetFormula(sheet, totalRow, 8, $"SUM({string.Join(",", categoryRows.Select(r => $"H{r}"))})");
            SetFormula(sheet, totalRow, 9, $"IFERROR((F{totalRow}+H{totalRow})/E{totalRow},0)");
            SetFormula(sheet, totalRow, 10, $"IFERROR(G{totalRow}/F{totalRow},0)");
            SetFormula(sheet, totalRow, 11, $"IFERROR((F{totalRow}-G{totalRow})/E{totalRow},0)");

            // 建立信息（历史报告里在标题下方）
            if (!string.IsNullOrWhiteSpace(request.CreatedBy))
            {
                ExcelNpoi.SetCell(sheet, 4, 4, $"Create by: {request.CreatedBy.Trim()}");
            }

            int lastTableRow = Math.Max(row - 1, firstDataRow);
            RewriteHyperlinks(sheet, targetFolder, itemNameRows);
            ApplyStatusBackground(wb, sheet, lastTableRow);
            // 表格边框与 Waterfall 一致：最外一圈粗黑线、内部细灰线（表头行 5 起、到最后一个测试项行为止）
            ExcelNpoi.ApplyBlockBorder(wb, sheet, 5, 2, lastTableRow, 12,
                BorderStyle.Thin, BorderStyle.Medium, IndexedGreyBorder, IndexedBlackBorder);
        }

        /// <summary>
        /// 按"测试种类"分组（按首次出现顺序；同类的测试项排在一起，与历史报告的 ORT Plan / TestStatus 一致）
        /// </summary>
        private static List<KeyValuePair<string, List<ReportTemplateItem>>> GroupByCategory(IEnumerable<ReportTemplateItem> items)
        {
            List<KeyValuePair<string, List<ReportTemplateItem>>> groups = [];
            foreach (ReportTemplateItem item in items ?? [])
            {
                string category = TestCategories.Normalize(item.Category) ?? TestCategories.Uncertain;
                int index = groups.FindIndex(g => string.Equals(g.Key, category, StringComparison.OrdinalIgnoreCase));
                if (index < 0)
                {
                    groups.Add(new KeyValuePair<string, List<ReportTemplateItem>>(category, [item]));
                }
                else
                {
                    groups[index].Value.Add(item);
                }
            }
            return groups;
        }

        /// <summary>取模板某一行的各列样式（下标 = 列号）</summary>
        private static ICellStyle[] CaptureRowStyles(ISheet sheet, int row1)
        {
            int lastCol = Math.Max(ExcelNpoi.LastColumn(sheet), 12);
            ICellStyle[] styles = new ICellStyle[lastCol + 1];
            for (int c = 1; c <= lastCol; c++)
            {
                styles[c] = ExcelNpoi.ExistingCell(sheet, row1, c)?.CellStyle;
            }
            return styles;
        }

        /// <summary>按列套用模板行样式（取不到样式的列保持原样）</summary>
        private static void ApplyRowStyles(ISheet sheet, int row1, ICellStyle[] styles)
        {
            if (styles == null)
            {
                return;
            }
            for (int c = 1; c < styles.Length; c++)
            {
                if (styles[c] != null)
                {
                    ExcelNpoi.Cell(sheet, row1, c).CellStyle = styles[c];
                }
            }
        }

        /// <summary>
        /// 重建测试项的超链接：每个测试项指向 Report 文件夹里它自己那份报告
        /// （历史报告里测试项名是蓝字带下划线，点了直接打开对应报告；模板里现成的链接行号会对不上）。
        /// </summary>
        private static void RewriteHyperlinks(ISheet sheet, string targetFolder, List<(int Row, string ItemName)> itemNameRows)
        {
            if (sheet is not XSSFSheet xssf)
            {
                return;
            }
            foreach (IHyperlink link in sheet.GetHyperlinkList().ToList())
            {
                xssf.RemoveHyperlink(link.FirstRow, link.FirstColumn);
            }
            List<string> reports = FindReportFiles(targetFolder);
            if (reports.Count == 0)
            {
                return;
            }
            foreach ((int row, string itemName) in itemNameRows)
            {
                string file = reports.FirstOrDefault(f => NameMatches(f, itemName));
                if (file == null)
                {
                    continue;
                }
                XSSFHyperlink link = new(HyperlinkType.File)
                {
                    Address = $"Report/{file}",
                    FirstRow = row - 1,
                    LastRow = row - 1,
                    FirstColumn = 3,
                    LastColumn = 3
                };
                xssf.AddHyperlink(link);
            }
        }

        /// <summary>报告文件夹（根目录下的 Report 子文件夹）里的报告文件</summary>
        private static List<string> FindReportFiles(string targetFolder)
        {
            try
            {
                string dir = string.IsNullOrWhiteSpace(targetFolder) ? null : Path.Combine(targetFolder, "Report");
                if (dir == null || !Directory.Exists(dir))
                {
                    return [];
                }
                return Directory.GetFiles(dir, "*.xls*")
                    .Select(Path.GetFileName)
                    .Where(f => !f.StartsWith("~$", StringComparison.Ordinal))
                    .OrderBy(f => f)
                    .ToList();
            }
            catch (Exception)
            {
                return [];
            }
        }

        /// <summary>报告文件名与测试项名是否对应（忽略大小写、空格、连字符与 "1.2 " 这类编号前缀）</summary>
        private static bool NameMatches(string fileName, string itemName)
        {
            string file = NormalizeForMatch(Path.GetFileNameWithoutExtension(fileName ?? ""));
            string item = NormalizeForMatch(itemName);
            if (file.Length == 0 || item.Length == 0)
            {
                return false;
            }
            // 去掉编号前缀（如 12ORTBURNINTESTREPORT → ORTBURNINTESTREPORT 不好剥，改用互相包含判断）
            return file.Contains(item) || item.Contains(file);
        }

        /// <summary>只留字母数字并转大写（"1.2 ORT Burn-In Test Report" → "12ORTBURNINTESTREPORT"）</summary>
        private static string NormalizeForMatch(string text)
        {
            StringBuilder builder = new();
            foreach (char ch in text ?? "")
            {
                if (char.IsLetterOrDigit(ch))
                {
                    builder.Append(char.ToUpperInvariant(ch));
                }
            }
            return builder.ToString();
        }

        /// <summary>
        /// 整表底色：全表铺淡灰（模板的列默认样式就是这个底，NPOI 写回会把"到最后一列"的默认样式丢掉，
        /// 所以这里显式铺出来），表格及表格周围一格刷白，已有的彩色底纹（黑表头、蓝色分类行等）保持不动。
        /// </summary>
        private static void ApplyStatusBackground(IWorkbook wb, ISheet sheet, int lastTableRow)
        {
            const int tableFirstCol = 2;
            const int tableLastCol = 12; // L 列（COMMENTS）
            // 铺得比表格大一圈，保证滚动到的区域都是淡灰而不是白
            int canvasLastRow = lastTableRow + 40;
            int canvasLastCol = tableLastCol + 28;
            // 先整体铺灰（连模板里原有的白底一起换掉），再把"表格 + 周围一格"刷白
            ExcelNpoi.ApplyFillOverlay(wb, sheet, 1, 1, canvasLastRow, canvasLastCol, ExcelNpoi.IndexedSilver, true);
            ExcelNpoi.ApplyFillOverlay(wb, sheet, 1, tableFirstCol - 1, lastTableRow + 1, tableLastCol + 1, ExcelNpoi.IndexedWhite, true);
        }

        /// <summary>
        /// 写公式：先清掉模板里残留的缓存值（否则会留着 #DIV/0! 之类的结果类型），再写公式并保留样式
        /// </summary>
        private static void SetFormula(ISheet sheet, int row, int col, string formula)
        {
            ICell cell = ExcelNpoi.Cell(sheet, row, col);
            ICellStyle style = cell.CellStyle;
            cell.SetCellValue((string)null);
            ExcelNpoi.SetFormula(sheet, row, col, formula);
            if (style != null)
            {
                cell.CellStyle = style;
            }
        }

        /// <summary>套用模板行样式（取不到样式时保持原有外观）</summary>
        private static void ApplyStyleOrKeep(ISheet sheet, int row1, int col1, int row2, int col2, ICellStyle style)
        {
            if (style != null)
            {
                ExcelNpoi.ApplyStyle(sheet, row1, col1, row2, col2, style);
            }
        }

        /// <summary>分类行的合计公式：区间为该分类下的测试项行</summary>
        private static void SumRow(ISheet sheet, int row, int firstItemRow, int lastItemRow)
        {
            SetFormula(sheet, row, 5, $"SUM(E{firstItemRow}:E{lastItemRow})");
            SetFormula(sheet, row, 6, $"SUM(F{firstItemRow}:F{lastItemRow})");
            SetFormula(sheet, row, 7, $"SUM(G{firstItemRow}:G{lastItemRow})");
            SetFormula(sheet, row, 8, $"SUM(H{firstItemRow}:H{lastItemRow})");
            SetFormula(sheet, row, 9, $"IFERROR((F{row}+H{row})/E{row},0)");
            SetFormula(sheet, row, 10, $"IFERROR(G{row}/F{row},0)");
            SetFormula(sheet, row, 11, $"IFERROR((F{row}-G{row})/E{row},0)");
        }

        /* ###############################  通用  ################################ */

        /// <summary>按名称查找工作表（忽略大小写与空格）</summary>
        private static ISheet FindSheet(IWorkbook wb, string keyword)
        {
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                string name = wb.GetSheetName(i);
                if (!string.IsNullOrWhiteSpace(name) && name.ToLowerInvariant().Replace(" ", "").Contains(keyword.Replace(" ", "")))
                {
                    return wb.GetSheetAt(i);
                }
            }
            return null;
        }

        /// <summary>移除与指定行区间相交的合并区域（重建表格前先清理）</summary>
        private static void RemoveMergesInRows(ISheet sheet, int row1, int row2)
        {
            int first = Math.Min(row1, row2) - 1;
            int last = Math.Max(row1, row2) - 1;
            for (int i = sheet.NumMergedRegions - 1; i >= 0; i--)
            {
                CellRangeAddress region = sheet.GetMergedRegion(i);
                if (region.LastRow >= first && region.FirstRow <= last)
                {
                    sheet.RemoveMergedRegion(i);
                }
            }
        }

        /// <summary>保证工作簿顺序为 Cover / ORT Plan / Waterfall / TestStatus</summary>
        private static void EnsureSheetOrder(IWorkbook wb)
        {
            string[] order = ["cover", "ort plan", "waterfall", "teststatus"];
            int position = 0;
            foreach (string keyword in order)
            {
                ISheet sheet = FindSheet(wb, keyword);
                if (sheet == null)
                {
                    continue;
                }
                int current = wb.GetSheetIndex(sheet);
                if (current != position)
                {
                    wb.SetSheetOrder(sheet.SheetName, position);
                }
                position++;
            }
            // TestStatus 名称统一（模板里可能是 "TestStatus for"）
            ISheet status = FindSheet(wb, "teststatus");
            if (status != null && !string.Equals(status.SheetName, "TestStatus", StringComparison.Ordinal))
            {
                wb.SetSheetName(wb.GetSheetIndex(status), "TestStatus");
            }
        }

        /* ###############################  由计划构建  ################################ */

        /// <summary>
        /// 用机种测试计划生成报告模板的输入（测试项按计划顺序，含共用模板里的公共文本）。
        /// 找不到计划时返回 null。
        /// </summary>
        public static ReportTemplateRequest BuildRequestFromPlan(TestPlan plan, List<TestPlanItem> items,
            string outputRoot, string createdBy)
        {
            if (plan == null)
            {
                return null;
            }
            ReportTemplateRequest request = new()
            {
                ModelName = plan.ModelName,
                Stage = plan.Stage,
                OutputRoot = outputRoot,
                CreatedBy = createdBy,
                PlanNote = plan.Remark
            };
            foreach (TestPlanItem item in items ?? [])
            {
                request.Items.Add(new ReportTemplateItem
                {
                    TestItemName = item.TestItemName,
                    Category = TestCategories.Normalize(item.Category ?? item.Template?.Category) ?? TestCategories.Uncertain,
                    SamplingPlan = item.EffectiveSamplingPlan,
                    TestCondition = item.EffectiveTestCondition,
                    PassCriterion = item.EffectivePassCriterion,
                    Remark = item.EffectiveRemark,
                    PeriodHours = item.EffectivePeriod
                });
            }
            return request;
        }
    }
}

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
                    WriteCover(wb, request);
                    WriteOrtPlan(wb, request);
                    WriteWaterfall(wb, request);
                    WriteTestStatus(wb, request);
                    EnsureSheetOrder(wb);
                    ExcelNpoi.Save(wb, overviewPath);
                }
                finally
                {
                    wb.Close();
                }

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

        /// <summary>递归复制目录（跳过 Excel 临时文件）</summary>
        private static void CopyDirectory(string source, string target)
        {
            Directory.CreateDirectory(target);
            foreach (string file in Directory.GetFiles(source))
            {
                if (Path.GetFileName(file).StartsWith("~$", StringComparison.Ordinal))
                {
                    continue;
                }
                File.Copy(file, Path.Combine(target, Path.GetFileName(file)), true);
            }
            foreach (string dir in Directory.GetDirectories(source))
            {
                CopyDirectory(dir, Path.Combine(target, Path.GetFileName(dir)));
            }
        }

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
        private static void WriteOrtPlan(IWorkbook wb, ReportTemplateRequest request)
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
            ExcelNpoi.ApplyStyle(sheet, 1, 4, 1, 4, new ExcelNpoi.CellStyleSpec { Bold = true, FontSize = 14 });
            ExcelNpoi.SetCell(sheet, 3, 2, string.IsNullOrWhiteSpace(request.CreatedBy) ? "" : $"Created By: {request.CreatedBy}");
            ExcelNpoi.SetCell(sheet, 3, 4, "Model: ");
            ExcelNpoi.SetCell(sheet, 3, 5, request.ModelName ?? "");

            string[] headers = ["NO.", "TEST\nITEMS", "SAMPLING PLAN", "TEST CONDITOIN", "PASS CRITERION", "COMMENT"];
            for (int i = 0; i < headers.Length; i++)
            {
                ExcelNpoi.SetCell(sheet, 4, 2 + i, headers[i]);
            }
            ExcelNpoi.ApplyStyle(sheet, 4, 2, 4, 7, new ExcelNpoi.CellStyleSpec
            {
                Bold = true,
                Border = BorderStyle.Thin,
                Horizontal = HorizontalAlignment.Center,
                Vertical = VerticalAlignment.Center,
                WrapText = true
            });

            int row = 6;
            int no = 0;
            int categoryNo = 0;
            string currentCategory = null;
            foreach (ReportTemplateItem item in request.Items ?? [])
            {
                string category = string.IsNullOrWhiteSpace(item.Category) ? currentCategory : item.Category.Trim();
                if (!string.IsNullOrWhiteSpace(category) && !string.Equals(category, currentCategory, StringComparison.OrdinalIgnoreCase))
                {
                    currentCategory = category;
                    categoryNo++;
                    no = 0;
                    ExcelNpoi.SetCell(sheet, row, 2, categoryNo);
                    ExcelNpoi.SetCell(sheet, row, 3, category);
                    ExcelNpoi.ApplyStyle(sheet, row, 2, row, 7, new ExcelNpoi.CellStyleSpec
                    {
                        Bold = true,
                        Border = BorderStyle.Thin,
                        Vertical = VerticalAlignment.Center
                    });
                    ExcelNpoi.Merge(sheet, row, 3, row, 7);
                    row++;
                }
                no++;
                ExcelNpoi.SetCell(sheet, row, 2, no);
                ExcelNpoi.SetCell(sheet, row, 3, item.TestItemName ?? "");
                ExcelNpoi.SetCell(sheet, row, 4, item.SamplingPlan ?? "");
                ExcelNpoi.SetCell(sheet, row, 5, item.TestCondition ?? "");
                ExcelNpoi.SetCell(sheet, row, 6, item.PassCriterion ?? "");
                ExcelNpoi.SetCell(sheet, row, 7, item.Remark ?? "");
                ExcelNpoi.ApplyStyle(sheet, row, 2, row, 7, new ExcelNpoi.CellStyleSpec
                {
                    Border = BorderStyle.Thin,
                    Vertical = VerticalAlignment.Top,
                    WrapText = true
                });
                row++;
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
        }

        /* ###############################  Waterfall  ################################ */

        /// <summary>
        /// Waterfall：D 列填序列号、序列号下一行填工令、E4 填起始日期，
        /// 右侧按排期写入每天的测试安排（同一天范围的同名测试合并单元格）。
        /// </summary>
        private static void WriteWaterfall(IWorkbook wb, ReportTemplateRequest request)
        {
            ISheet sheet = FindSheet(wb, "waterfall") ?? ExcelNpoi.SheetAt(wb, 2);
            if (sheet == null)
            {
                return;
            }
            const int firstSnRow = 7;
            const int dateRow = 4;
            const int firstDateCol = 5; // E 列
            List<string> sns = (request.SerialNumbers ?? []).Where(s => !string.IsNullOrWhiteSpace(s)).Select(s => s.Trim()).ToList();
            int snCount = Math.Max(sns.Count, 1);

            // 1. 序列号行数按实际调整（模板里预留了 3 行）
            const int templateSnRows = 3;
            if (snCount < templateSnRows)
            {
                ExcelNpoi.DeleteRows(sheet, firstSnRow + snCount, templateSnRows - snCount);
            }
            else if (snCount > templateSnRows)
            {
                int extra = snCount - templateSnRows;
                int totalCols = Math.Max(ExcelNpoi.LastColumn(sheet), firstDateCol);
                for (int i = 0; i < extra; i++)
                {
                    int insertAt = firstSnRow + templateSnRows + i;
                    ExcelNpoi.InsertRows(sheet, insertAt, 1);
                    ExcelNpoi.CopyRowStyle(sheet, firstSnRow + templateSnRows - 1, insertAt, totalCols);
                }
            }

            // 2. 清空并重建"测试安排"区域（先去掉该区域的合并，避免残留）；
            //    B 列是单体序号（模板里 1、2、3），清空时保留并重写
            int lastRow = ExcelNpoi.LastRow(sheet);
            int clearLastRow = firstSnRow + snCount + 1; // 含工令行
            int lastCol = Math.Max(ExcelNpoi.LastColumn(sheet), firstDateCol + snCount * 20);
            RemoveMergesInRows(sheet, firstSnRow, Math.Max(clearLastRow, lastRow));
            for (int r = firstSnRow; r <= clearLastRow; r++)
            {
                ExcelNpoi.ClearCells(sheet, r, 3, lastCol);
                if (r < firstSnRow + snCount)
                {
                    ExcelNpoi.SetCell(sheet, r, 2, r - firstSnRow + 1);
                }
            }

            // 3. 起始日期 + 序列号 + 工令
            //    起始日期对齐到工作日：第一项测试必然落在 E 列，概览读取（Waterfall 的 S/N 判定）也才稳
            DateTime start = NextWorkingDay(request.StartDate);
            ExcelNpoi.SetCell(sheet, dateRow, firstDateCol, start);
            for (int i = 0; i < sns.Count; i++)
            {
                ExcelNpoi.SetCell(sheet, firstSnRow + i, 4, sns[i]);
            }
            if (!string.IsNullOrWhiteSpace(request.WorkOrder))
            {
                ExcelNpoi.SetCell(sheet, firstSnRow + snCount, 4, request.WorkOrder.Trim());
            }

            // 4. 每一天的日期/星期/日号（超出模板已有列时补公式与样式）
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
            int existingLastCol = ExcelNpoi.LastColumn(sheet);
            for (int c = firstDateCol; c <= lastDateCol; c++)
            {
                if (c == firstDateCol)
                {
                    ExcelNpoi.SetCell(sheet, dateRow, c, start);
                }
                else
                {
                    ICell cell = ExcelNpoi.Cell(sheet, dateRow, c);
                    cell.CellFormula = $"{ExcelNpoi.AddressOf(dateRow, c - 1)}+1";
                    cell.CellStyle = ExcelNpoi.DateStyleOf(wb);
                    ICell weekCell = ExcelNpoi.Cell(sheet, dateRow - 1, c);
                    weekCell.CellFormula = $"WEEKDAY({ExcelNpoi.AddressOf(dateRow, c)},2)";
                    ICell dayCell = ExcelNpoi.Cell(sheet, dateRow + 1, c);
                    dayCell.CellFormula = $"DAY({ExcelNpoi.AddressOf(dateRow, c)})";
                    // 样式沿用前一列的（模板列已有日期/星期/日号的格式）
                    for (int r = dateRow - 1; r <= dateRow + 1; r++)
                    {
                        ICellStyle style = ExcelNpoi.ExistingCell(sheet, r, c - 1)?.CellStyle;
                        if (style != null)
                        {
                            ExcelNpoi.Cell(sheet, r, c).CellStyle = style;
                        }
                    }
                }
            }
            for (int c = existingLastCol + 1; c <= lastDateCol; c++)
            {
                ExcelNpoi.SetColumnWidth(sheet, c, 5.4);
            }

            // 5. 测试安排：每个序列号都走一遍计划里的测试（同一天范围合并）
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
                for (int i = 0; i < snCount; i++)
                {
                    int row = firstSnRow + i;
                    ExcelNpoi.SetCell(sheet, row, from, item.TestItemName ?? "");
                    ExcelNpoi.ApplyStyle(sheet, row, from, row, to, new ExcelNpoi.CellStyleSpec
                    {
                        Horizontal = HorizontalAlignment.Center,
                        Vertical = VerticalAlignment.Center,
                        WrapText = true
                    });
                    if (to > from)
                    {
                        ExcelNpoi.Merge(sheet, row, from, row, to);
                    }
                }
            }
        }

        /* ###############################  TestStatus  ################################ */

        /// <summary>
        /// TestStatus：按计划重建"分类行 + 测试项行"，并重算合计行。
        /// 模板的数据行先整段删除再按需要重建，避免残留模板里的旧测试项；
        /// 分类行/测试项行的样式取模板既有行，保持外观一致。
        /// </summary>
        private static void WriteTestStatus(IWorkbook wb, ReportTemplateRequest request)
        {
            ISheet sheet = FindSheet(wb, "teststatus") ?? FindSheet(wb, "test status");
            if (sheet == null)
            {
                return;
            }
            const int totalRow = 8;
            const int firstDataRow = 9;
            // 模板第 9 行是分类行、第 10 行是测试项行，先把样式留下来
            ICellStyle categoryStyle = ExcelNpoi.ExistingCell(sheet, firstDataRow, 2)?.CellStyle;
            ICellStyle itemStyle = ExcelNpoi.ExistingCell(sheet, firstDataRow + 1, 2)?.CellStyle;
            int templateLast = ExcelNpoi.LastRow(sheet);
            int templateRows = Math.Max(templateLast - firstDataRow + 1, 0);

            // 用列表保留分类的先后顺序
            List<KeyValuePair<string, List<ReportTemplateItem>>> groups = [];
            foreach (ReportTemplateItem item in request.Items ?? [])
            {
                string category = string.IsNullOrWhiteSpace(item.Category) ? "RELIABILITY TEST" : item.Category.Trim();
                category = TestStatusCategory(category);
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
            ApplyStyleOrKeep(sheet, totalRow, 2, totalRow, 12, categoryStyle);

            int row = firstDataRow;
            int categoryNo = 0;
            List<int> categoryRows = [];
            foreach (KeyValuePair<string, List<ReportTemplateItem>> group in groups)
            {
                categoryNo++;
                categoryRows.Add(row);
                int firstItemRow = row + 1;
                int lastItemRow = row + group.Value.Count;
                ExcelNpoi.SetCell(sheet, row, 2, categoryNo);
                ExcelNpoi.SetCell(sheet, row, 4, group.Key);
                ExcelNpoi.Merge(sheet, row, 2, row, 3);
                ApplyStyleOrKeep(sheet, row, 2, row, 12, categoryStyle);
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
                    ApplyStyleOrKeep(sheet, row, 2, row, 12, itemStyle);
                    // 完成率/% 与合计行同构：由已录入的完成数自动计算
                    SetFormula(sheet, row, 9, $"IFERROR((F{row}+H{row})/E{row},0)");
                    SetFormula(sheet, row, 10, $"IFERROR(G{row}/F{row},0)");
                    SetFormula(sheet, row, 11, $"IFERROR((F{row}-G{row})/E{row},0)");
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
        }

        /// <summary>
        /// TestStatus 的分类名：既有报告里环境类测试写的是 ENVIRONMENT TESTS、EMC 组写的是 "EMC "，
        /// 这里按同样写法转换，保持与历史报告一致（ORT Plan 表仍用计划里的 RELIABILITY TEST）。
        /// </summary>
        private static string TestStatusCategory(string category)
        {
            string text = (category ?? "").Trim();
            if (text.Equals("RELIABILITY TEST", StringComparison.OrdinalIgnoreCase)
                || text.Equals("ENVIRONMENT TEST", StringComparison.OrdinalIgnoreCase))
            {
                return "ENVIRONMENT TESTS";
            }
            if (text.Equals("EMC", StringComparison.OrdinalIgnoreCase))
            {
                return "EMC ";
            }
            return text;
        }

        /// <summary>
        /// 写公式：先清掉模板里残留的缓存值（否则会留着 #DIV/0! 之类的结果类型），再写公式并保留样式
        /// </summary>
        private static void SetFormula(ISheet sheet, int row, int col, string formula)
        {
            ICell cell = ExcelNpoi.Cell(sheet, row, col);
            ICellStyle style = cell.CellStyle;
            cell.SetCellValue((string)null);
            cell.CellFormula = formula;
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
                    Category = item.Category ?? item.Template?.Category,
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

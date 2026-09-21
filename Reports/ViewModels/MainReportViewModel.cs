using CommunityToolkit.Mvvm.ComponentModel;
using NLog;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using ExcelNpoi = ORT一键报告.Utils.ExcelNpoi;
using ORT一键报告.Models;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using static ORT一键报告.Utils.Report;

namespace ORT一键报告.Reports.ViewModels
{
    public partial class MainReportViewModel(IPathService service, ReportService reportService, DatabaseService databaseService, AppSettingsService appSettings) : ObservableObject
    {
        private readonly IPathService Service = service;
        private readonly ReportService _reportService = reportService;
        private readonly DatabaseService _databaseService = databaseService;
        private readonly AppSettingsService _appSettings = appSettings;
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        public string ATEPath { get; set; }


        private string _reportPath;
        public string ReportPath
        {
            get => _reportPath;
            set
            {
                if (SetProperty(ref _reportPath, value))
                {
                    _reportService.RootPath = Path.GetDirectoryName(value);
                    selectReportPathCommand.RaiseCanExecuteChanged();
                }
            }
        }

        private string _title = LanguageService.Get("App_ReportTitle");
        public string Title
        {
            get => _title;
            set => SetProperty(ref _title, value);
        }

        private RelayCommand selectReportPathCommand;
        public ICommand SelectReportPathCommand => selectReportPathCommand ??= new RelayCommand(SelectReportPath);


        private void SelectReportPath()
        {
            ReportPath = Service.OpenPathDialog(LanguageService.Get("Dlg_SelectReportOverview"), initPath: _appSettings.ReportDir);
            if (ReportPath == null)
            {
                return;
            }
            string _title = Path.GetFileName(Path.GetDirectoryName(ReportPath));
            string productName = LanguageService.Get("App_ReportTitle");
            try
            {
                Title = _title.Split(' ')[0] + " " + _title.Split('_')[1] + " " + productName;
            }
            catch
            {
                Title = " " + productName;
            }
            // 数据源重构：按文件夹名称中的机种名称、RT工作编号等信息从领退和计划中匹配记录
            MatchPlanFromFolderName(_title);
        }

        /// <summary>
        /// 从报告文件夹名称中解析 RT工号/机种名称，在领退和计划中匹配记录（优先RT工号，其次机种），
        /// 匹配结果存入 ReportService.MatchedPlan，供报告表头补充信息。
        /// </summary>
        private void MatchPlanFromFolderName(string folderName)
        {
            // 保留从计划表右键携带的预填记录：文件夹名未匹配到时仍使用携带记录
            Plan prefilled = _reportService.MatchedPlan;
            _reportService.MatchedPlan = null;
            if (string.IsNullOrWhiteSpace(folderName))
            {
                return;
            }
            try
            {
                Plan matched = null;
                // 1. RT工号/回线工令（如 RT260637 / RTAH2610002）
                Match rt = Regex.Match(folderName, @"RT[A-Z]*\d+");
                if (rt.Success)
                {
                    string rtValue = rt.Value;
                    matched = _databaseService.FreeSql.Select<Plan>()
                        .Where(p => p.JobNo == rtValue)
                        .First();
                    if (matched == null)
                    {
                        // 回线工令属于领退表，找到后按机种匹配计划表
                        Requisition req = _databaseService.FreeSql.Select<Requisition>()
                            .Where(r => r.ReturnRtOrder == rtValue || r.WorkOrder == rtValue)
                            .First();
                        if (req != null)
                        {
                            matched = _databaseService.FreeSql.Select<Plan>()
                                .Where(p => p.ModelName == req.ModelName)
                                .OrderByDescending(p => p.Id)
                                .First();
                        }
                    }
                }
                // 2. 机种名称（如 FSA037-4B1G，通常位于文件夹名开头）
                if (matched == null)
                {
                    Match model = Regex.Match(folderName, @"^([A-Za-z]{2,5}\d{2,5}-[A-Za-z0-9]+)");
                    if (model.Success)
                    {
                        string modelName = model.Groups[1].Value.ToUpper();
                        List<Plan> candidates = _databaseService.FreeSql.Select<Plan>()
                            .Where(p => p.ModelName != null)
                            .ToList();
                        matched = candidates.FirstOrDefault(p => p.ModelName?.ToUpper() == modelName);
                    }
                }
                _reportService.MatchedPlan = matched ?? prefilled;
                if (_reportService.MatchedPlan != null)
                {
                    _logger.Info($"报告数据匹配到计划记录: Id={_reportService.MatchedPlan.Id} 机种={_reportService.MatchedPlan.ModelName} 工作編號={_reportService.MatchedPlan.JobNo} 负责人={_reportService.MatchedPlan.Owner}");
                }
                else
                {
                    _logger.Warn($"文件夹[{folderName}]未匹配到领退和计划中的记录，报告表头将仅使用模板信息");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "从领退和计划匹配记录失败");
            }
        }

        public async Task ReadInfoFromOverview(string ReportName)
        {
            _logger.Info("读取报告概览...");

            DateTime t_start = DateTime.Now;
            DateTime b_start = DateTime.Now;
            try
            {
                _reportService.UUTInfos = await Task.Run(() =>
                {
                    // 概览可能是旧的 .xls，按内容选引擎打开（原来只用 XSSF，打开 .xls 会抛 SharpZipLib "EOF in header"）
                    IWorkbook wb = ExcelNpoi.OpenAny(ReportName);
                    try
                    {
                        return ReadInfosFromReport(wb, ReportName);
                    }
                    finally
                    {
                        wb.Close();
                    }
                });
                // 从计划表右键进入时：SN/工令/版本/DC 以领用表为准（报告概览只用来补充测试项目等）
                if (_reportService.ApplyMatchedSourceToUUTInfos())
                {
                    _logger.Info("已用领用表数据覆盖报告概览中的 SN/工令/版本/DC");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "读取报告概览时出现错误");
                return;
            }
            if (_reportService.UUTInfos == null)
            {
                _logger.Warn("读取报告概览返回空数据，跳过测试项日期解析");
                return;
            }
            // 概览里没读到序列号时明确告知用户（并已在日志里记录工作表/标题定位情况），不再静默无 SN
            if ((_reportService.UUTInfos.SNs?.Count ?? 0) == 0)
            {
                _logger.Warn("报告概览里没有读到序列号");
                _ = MessageBox.Show(
                    "没能从报告概览里读到序列号（Waterfall 表的 S/N 列）。\n"
                    + "请确认选择的是该次测试的报告概览文件；详细原因见日志。",
                    LanguageService.Get("Cap_Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            foreach (TestItemInfo testItem in _reportService.UUTInfos.TestItems ?? [])
            {
                string name = testItem.TestItemName?.ToLower() ?? "";
                if (name.Contains("thermal shock") || name.Contains("burn in"))
                {
                    if (DateTime.TryParse(testItem.Date, out DateTime parsedDate))
                    {
                        t_start = parsedDate;
                    }
                    else
                    {
                        _logger.Warn($"测试项目 {testItem.TestItemName} 的日期无效（{testItem.Date}），跳过日期解析");
                    }
                }
            }

            UUTInfoFromExcel ReadInfosFromReport(IWorkbook wb, string _ReportName)
            {
                // 概览的工作表顺序不固定：优先按名称找 Cover / Waterfall，找不到再退回原来的位置
                ISheet ws_cover = FindSheetByName(wb, "cover") ?? ExcelNpoi.SheetAt(wb, 0);
                ISheet ws_waterfall = FindSheetByName(wb, "waterfall") ?? ExcelNpoi.SheetAt(wb, 2);
                _logger.Info("报告概览工作表：" + string.Join(" | ", Enumerable.Range(0, wb.NumberOfSheets).Select(i => wb.GetSheetName(i)))
                    + $"；使用 Cover='{ws_cover?.SheetName}' Waterfall='{ws_waterfall?.SheetName}'");
                if (ws_cover == null || ws_waterfall == null)
                {
                    _logger.Error($"报告概览缺少工作表：cover={ws_cover != null} waterfall={ws_waterfall != null}");
                    return null;
                }
                // 测试周期取自文件夹名里的 WK####（取不到再看文件名），ISO 周号 -> 当周周一；DC 不再从这个编号取
                string folderName = Path.GetDirectoryName(_ReportName) is string dir && dir.Length > 0
                    ? Path.GetFileName(dir)
                    : null;
                DateTime? weekPeriod = ParseWeekPeriod(folderName) ?? ParseWeekPeriod(_ReportName);
                UUTInfoFromExcel uutInfos = new()
                {
                    TestPeriod = ParseWeekTag(folderName) ?? ParseWeekTag(_ReportName),
                    TestStart = weekPeriod
                };
                if (weekPeriod != null)
                {
                    _logger.Info($"从报告文件夹名解析到测试周期：{uutInfos.TestPeriod} -> {weekPeriod.Value:yyyy-MM-dd}（当周周一）");
                }

                DataCell rev = FindCellByValue(ws_cover, "rev");
                if (rev == null)
                {
                    // 注意：这里是后台线程，不能弹 MessageBox（会阻塞任务让上层永远等不到结果）
                    _logger.Warn("报告概览：Cover 表里没找到 Revision 标签，版本留空");
                }
                else
                {
                    for (int c = rev.Column + 1; c < ExcelNpoi.LastColumn(ws_cover); c++)
                    {
                        if (ExcelNpoi.CellText(ws_cover, rev.Row, c) != "")
                        {
                            uutInfos.Revision = ExcelNpoi.CellText(ws_cover, rev.Row, c);
                        }
                    }
                }

                // S/N 标题：先按"含 s/n 但不含 uut"，再放宽为含 s/n / serial（不同版本概览表头写法有差异）
                DataCell snTitleCell = FindCellByValue(ws_waterfall, "s/n", "uut")
                    ?? FindCellByValue(ws_waterfall, "s/n")
                    ?? FindCellByValue(ws_waterfall, "serial");
                if (snTitleCell == null)
                {
                    _logger.Warn($"报告概览：Waterfall({ws_waterfall.SheetName}) 表里没找到 S/N 标题，序列号留空");
                    return uutInfos;
                }

                List<DataCell> snCells = FindSNs(ws_waterfall, snTitleCell);
                if (snCells.Count == 0)
                {
                    _logger.Warn($"报告概览：S/N 标题在 {ExcelNpoi.AddressOf(snTitleCell.Row, snTitleCell.Column)}，但其下方没有找到序列号");
                    return uutInfos;
                }
                List<string> SNs = [];
                foreach (DataCell cell in snCells)
                {
                    SNs.Add(cell.Data);
                }
                uutInfos.SNs = SNs;
                uutInfos.WorkOrder = ExcelNpoi.CellText(ws_waterfall, snCells.Last().Row + 1, snCells.Last().Column);
                List<TestItemInfo> TestItems = FindTestItems(ws_waterfall, snTitleCell.Row, snCells.First().Row, snCells.First().Column);
                uutInfos.TestItems = TestItems;
                _logger.Info($"报告概览读取完成：SN {SNs.Count} 个、工令='{uutInfos.WorkOrder}'、版本='{uutInfos.Revision}'、周期={uutInfos.TestPeriod}");
                return uutInfos;
            }

            static ISheet FindSheetByName(IWorkbook workbook, string keyword)
            {
                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    string name = workbook.GetSheetName(i);
                    if (!string.IsNullOrWhiteSpace(name) && name.ToLowerInvariant().Contains(keyword))
                    {
                        return workbook.GetSheetAt(i);
                    }
                }
                return null;
            }

            List<TestItemInfo> FindTestItems(ISheet ws, int rDate, int rSN, int cSN)
            {
                List<TestItemInfo> testItems = [];
                int c = cSN + 1;
                for (; c <= ExcelNpoi.LastColumn(ws); c++)
                {
                    if (ExcelNpoi.CellText(ws, rSN, c) is string testitem && testitem != "")
                    {
                        string date = ExcelNpoi.CellText(ws, rDate, c);
                        testItems.Add(new TestItemInfo
                        {
                            TestItemName = testitem,
                            Date = date
                        });
                    }
                }
                return testItems;
            }

            List<DataCell> FindSNs(ISheet ws, DataCell snTitleCell)
            {
                /// <summary>
                /// 在指定范围内寻找单元格值为"S/N"的单元格，找到后继续向下寻找非空且右边也非空的单元格，直到遇到空单元格为止，将这些非空单元格的信息（值、行号、列号）存储在SNCell对象中，并返回一个包含所有SNCell对象的列表。
                /// </summary>
                List<DataCell> snCells = [];
                int rSN = snTitleCell.Row + 1;
                int cSN = snTitleCell.Column;
                for (; rSN <= ExcelNpoi.LastRow(ws); rSN++)
                {
                    if (ExcelNpoi.CellText(ws, rSN, cSN) is string sn && sn != "")
                    {
                        if (ExcelNpoi.CellText(ws, rSN, cSN + 1) is "")
                        {
                            continue;
                        }
                        snCells.Add(new DataCell(rSN, cSN) { Data = sn });
                    }
                }
                return snCells;
            }
        }
    }
}

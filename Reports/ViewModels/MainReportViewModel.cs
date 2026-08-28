using CommunityToolkit.Mvvm.ComponentModel;
using NLog;
using OfficeOpenXml;
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

        private string _title = "ORT一键报告";
        public string Title
        {
            get => _title;
            set => SetProperty(ref _title, value);
        }

        private RelayCommand selectReportPathCommand;
        public ICommand SelectReportPathCommand => selectReportPathCommand ??= new RelayCommand(SelectReportPath);


        private void SelectReportPath()
        {
            ReportPath = Service.OpenPathDialog("选择报告概览", initPath: _appSettings.ReportDir);
            if (ReportPath == null)
            {
                return;
            }
            string _title = Path.GetFileName(Path.GetDirectoryName(ReportPath));
            try
            {
                Title = _title.Split(' ')[0] + " " + _title.Split('_')[1] + " ORT一键报告";
            }
            catch
            {
                Title = " ORT一键报告";
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
                    ExcelPackage package = new(new FileInfo(ReportName));
                    ExcelWorkbook wb = package.Workbook;
                    return ReadInfosFromReport(wb, ReportName);
                });
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "读取报告概览时出现错误");
                return;
            }
            foreach (TestItemInfo testItem in _reportService.UUTInfos.TestItems)
            {
                if (testItem.TestItemName.ToLower().Contains("thermal shock"))
                {
                    t_start = DateTime.Parse(testItem.Date);
                }
                if (testItem.TestItemName.ToLower().Contains("burn in"))
                {
                    t_start = DateTime.Parse(testItem.Date);
                }
            }

            UUTInfoFromExcel ReadInfosFromReport(ExcelWorkbook wb, string _ReportName)
            {
                var ws_cover = wb.Worksheets[0];
                var ws_waterfall = wb.Worksheets[2];
                UUTInfoFromExcel uutInfos = new()
                {
                    DC = GetSubstringAfter(_ReportName, "WK", 4)
                };

                DataCell rev = FindCellByValue(ws_cover, "rev");
                if (rev == null)
                {
                    MessageBox.Show("未找到Rev列", "错误");
                    return null;
                }
                for (int c = rev.Column + 1; c < ws_cover.Dimension.End.Column; c++)
                {
                    if (ws_cover.Cells[rev.Row, c].Text != "")
                    {
                        uutInfos.Revision = ws_cover.Cells[rev.Row, c].Text;
                    }
                }

                DataCell snTitleCell = FindCellByValue(ws_waterfall, "s/n", "uut");
                if (snTitleCell == null)
                {
                    MessageBox.Show("未找到SN列", "错误");
                    return null;
                }

                List<DataCell> snCells = FindSNs(ws_waterfall, snTitleCell);
                if (snCells.Count == 0)
                {
                    MessageBox.Show("没有SN", "错误");
                    return null;
                }
                else
                {
                    List<string> SNs = [];
                    foreach (DataCell cell in snCells)
                    {
                        SNs.Add(cell.Data);
                    }
                    uutInfos.SNs = SNs;
                    uutInfos.WorkOrder = ws_waterfall.Cells[snCells.Last().Row + 1, snCells.Last().Column].Text;
                }
                List<TestItemInfo> TestItems = FindTestItems(ws_waterfall, snTitleCell.Row, snCells.First().Row, snCells.First().Column);
                uutInfos.TestItems = TestItems;

                return uutInfos;
            }

            List<TestItemInfo> FindTestItems(ExcelWorksheet ws, int rDate, int rSN, int cSN)
            {
                List<TestItemInfo> testItems = [];
                int c = cSN + 1;
                for (; c <= ws.Dimension.End.Column; c++)
                {
                    if (ws.Cells[rSN, c].Text is string testitem && testitem != "")
                    {
                        string date = ws.Cells[rDate, c].Text;
                        testItems.Add(new TestItemInfo
                        {
                            TestItemName = testitem,
                            Date = date
                        });
                    }
                }
                return testItems;
            }

            List<DataCell> FindSNs(ExcelWorksheet ws, DataCell snTitleCell)
            {
                /// <summary>
                /// 在指定范围内寻找单元格值为"S/N"的单元格，找到后继续向下寻找非空且右边也非空的单元格，直到遇到空单元格为止，将这些非空单元格的信息（值、行号、列号）存储在SNCell对象中，并返回一个包含所有SNCell对象的列表。
                /// </summary>
                List<DataCell> snCells = [];
                int rSN = snTitleCell.Row + 1;
                int cSN = snTitleCell.Column;
                for (; rSN <= ws.Dimension.End.Row; rSN++)
                {
                    if (ws.Cells[rSN, cSN].Text is string sn && sn != "")
                    {
                        if (ws.Cells[rSN, cSN + 1].Text is "")
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

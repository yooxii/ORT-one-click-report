using NLog;
using OfficeOpenXml;
using ORT一键报告.Models;
using ORT一键报告.ViewModels;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using static ORT一键报告.ReportUtils;

namespace ORT一键报告
{
    public abstract class DynamicField
    {
        public string Label { get; set; } // 控件左侧的标签文字
    }

    public class TextField : DynamicField
    {
        public string Value { get; set; }
    }

    public class SwitchField : DynamicField
    {
        public bool IsChecked { get; set; }
    }

    public class OptionField : DynamicField
    {
        public List<string> Options { get; set; }
        public string SelectedOption { get; set; }
    }

    public enum ReportStatus
    { Pass, Fail };

    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public MainViewModel MainVM { get; set; }

        public static SettingsViewModel SettingsVM { get; set; }

        public static UUTInfoFromExcel UUTInfos { get; set; }

        public static string RootPath { get; set; }
        public static string TemplateDir { get; set; }
        public static string TempPath { get; set; }

        private readonly Dictionary<string, object> defaultSetup = new() { 
            {"路径对话框初始目录", new Dictionary<string, object> { 
                {"BI EMI 报告","\\\\bnt56\\品保部\\ORT實驗資料\\13. 臨時試驗報告\\BI EMI"},
                {"BI ATE Data", "\\\\bnt56\\品保部\\ORT實驗資料\\13. 臨時試驗報告\\BI ATE Data" },
                {"BI Picture","\\\\bnt56\\品保部\\ORT實驗資料\\13. 臨時試驗報告\\BI Picture" }
            } },
        };

        public MainWindow()
        {
            InitializeComponent();
            ExcelPackage.License.SetNonCommercialPersonal("Lucas");

            SettingsVM = new SettingsViewModel();
            MainVM = new(new Service());
            DataContext = MainVM;

            Closed += Window_Closed;
            Loaded += ReportHeader_Loaded;

            TemplateDir = Path.Combine(Directory.GetCurrentDirectory(), "Templates");
            TempPath = Path.Combine(Path.GetTempPath(), "ORTTemp");
        }

        private void ReportHeader_Loaded(object sender, RoutedEventArgs e)
        {
            thermalshockPage.InitReportPage();
            burninPage.InitReportPage();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            ClearTempDir();
        }

        /* ###############################  功能函数  ################################ */

        private async Task ReadInfoFromOverview(string ReportName)
        {
            _logger.Info("读取报告概览...");

            DateTime t_start = DateTime.Now;
            DateTime b_start = DateTime.Now;
            try
            {
                UUTInfos = await Task.Run(() =>
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
            foreach (TestItemInfo testItem in UUTInfos.TestItems)
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
                    List<string> SNs = new();
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
                List<TestItemInfo> testItems = new();
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
                List<DataCell> snCells = new();
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
                        snCells.Add(new DataCell
                        {
                            Data = sn,
                            Row = rSN,
                            Column = cSN
                        });
                    }
                }
                return snCells;
            }
        }

        /* ###############################  事件函数  ################################ */

        private void MenuItem_ATE_Click(object sender, RoutedEventArgs e)
        {
            ATEWindow ateWindow = new();
            ateWindow.Show();
        }

        private void MenuItem_MainSetup_Click(object sender, RoutedEventArgs e)
        {

        }

        private async void DoReport_Click(object sender, RoutedEventArgs e)
        {
            PopupWindow popup = new() { Title = "处理中", Message = "请耐心等待..." };
            Button btn;
            if (sender is Button tmp)
            {
                btn = tmp;
            }
            else
            {
                return;
            }
            btn.IsEnabled = false;

            try
            {
                popup.Show();
                string ReportName = MainVM.ReportPath;
                if (!File.Exists(ReportName))
                {
                    throw new FileNotFoundException("报告概览文件不存在");
                }
                await ReadInfoFromOverview(ReportName);
                _logger.Info("报告概览读取完成");

                thermalshockPage.ReadReportHeader();
                thermalshockPage.SetReportResultData();
                burninPage.ReadReportHeader();
                burninPage.SetReportResultData();
                emiPage.ReadReportHeader();
                _logger.Info("表头数据已呈现至窗口");
            }
            catch (FileNotFoundException ex)
            {
                _logger.Error(ex, "报告文件不存在");
                _ = MessageBox.Show("报告文件不存在, 请正确选择", "错误");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "读取报告出现错误");
                _ = MessageBox.Show($"读取报告出现错误{ex}", "错误");
            }
            finally
            {
                popup.Close();
                btn.IsEnabled = true;
            }
        }
    }
}
using Microsoft.Win32;
using NLog;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using static ORT一键报告.ReportUtils;

namespace ORT一键报告
{
    public enum ReportStatus { Pass, Fail };

    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public ReportHeaderInfo burnReportHeaderInfo = null;
        public ReportHeaderInfo emiReportHeaderInfo = null;
        public ObservableCollection<ResultDetails> DetailsList = new ObservableCollection<ResultDetails>();

        public static UUTInfoFromExcel UUTInfos;
        public static string ATEPath { get; set; }
        public static string RootPath { get; set; }
        public static string TempPath = Path.Combine(Path.GetTempPath(), "ORTTemp");

        public MainWindow()
        {
            InitializeComponent();
            ExcelPackage.License.SetNonCommercialPersonal("Lucas");
            Closed += Window_Closed;
            Loaded += ReportHeader_Loaded;
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
            string SNsCount = "3";
            try
            {
                UUTInfos = await Task.Run(() =>
                {
                    var fileInfo = new FileInfo(ReportName);
                    var package = new ExcelPackage(fileInfo);
                    var wb = package.Workbook;
                    return ReadInfosFromReport(wb, ReportName);
                });
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "读取报告概览时出现错误");
                return;
            }
            DetailsList.Clear();
            foreach (var testItem in UUTInfos.TestItems)
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
            SNsCount = UUTInfos.SNs.Count.ToString();
            foreach (var uutInfo in UUTInfos.SNs)
            {
                DetailsList.Add(new ResultDetails()
                {
                    BIroom = "1F Chamber",
                    SN = uutInfo,
                    WorkOrder = UUTInfos.WorkerNo,
                    Version = UUTInfos.Revision,
                    DC = UUTInfos.DC,
                    InspectionPrev = ReportStatus.Pass,
                    FunPrev = ReportStatus.Pass,
                    InspectionAfter = ReportStatus.Pass,
                    FunAfter = ReportStatus.Pass,
                    HiPot = ReportStatus.Pass,
                    Comments = ""
                });
            }

            UUTInfoFromExcel ReadInfosFromReport(ExcelWorkbook wb, string _ReportName)
            {
                var ws_cover = wb.Worksheets[0];
                var ws_waterfall = wb.Worksheets[2];
                UUTInfoFromExcel uutInfos = new UUTInfoFromExcel
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
                    List<string> SNs = new List<string>();
                    foreach (var cell in snCells)
                    {
                        SNs.Add(cell.Data);
                    }
                    uutInfos.SNs = SNs;
                    uutInfos.WorkerNo = ws_waterfall.Cells[snCells.Last().Row + 1, snCells.Last().Column].Text;
                }
                List<TestItemInfo> TestItems = FindTestItems(ws_waterfall, snTitleCell.Row, snCells.First().Row, snCells.First().Column);
                uutInfos.TestItems = TestItems;

                return uutInfos;
            }

            List<TestItemInfo> FindTestItems(ExcelWorksheet ws, int rDate, int rSN, int cSN)
            {
                List<TestItemInfo> testItems = new List<TestItemInfo>();
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
                List<DataCell> snCells = new List<DataCell>();
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

        private void SetInfoToWindow()
        {
            fun(e_ReportHeader, emiReportHeaderInfo, widget_pic_e);

            void SetPics(List<ExcelPictureInfo> _pics, List<Image> images)
            {
                for (int i = 0; i < _pics.Count && i < 3; i++)
                {
                    images[i].Source = _pics[i].ImageSrc;
                }
            }

            void fun(ReportHeaderWidget ReportHeader, ReportHeaderInfo reportHeaderInfo, ReportPicturesWidget reportPicturesWidget)
            {
                ReportHeader.ApprovedBy = reportHeaderInfo.APPROVED_BY.Data;
                ReportHeader.TestedBy = reportHeaderInfo.TESTED_BY.Data;
                ReportHeader.ProjectName = reportHeaderInfo.PROJECT_NAME.Data;
                ReportHeader.TestStage = reportHeaderInfo.TEST_STAGE.Data;
                ReportHeader.text_TestDescription.Text = reportHeaderInfo.TestDescription.Data;

                if (reportHeaderInfo.Issue_Photos_Pics != null)
                {
                    SetPics(reportHeaderInfo.Issue_Photos_Pics.Images, new List<Image> { reportPicturesWidget.issue_image1, reportPicturesWidget.issue_image2, reportPicturesWidget.issue_image3 });
                }
                if (reportHeaderInfo.Test_Setup_Pics != null)
                {
                    SetPics(reportHeaderInfo.Test_Setup_Pics.Images, new List<Image> { reportPicturesWidget.setup_image1, reportPicturesWidget.setup_image2, reportPicturesWidget.setup_image3 });
                }
            }
        }

        /* ###############################  事件函数  ################################ */
        private void MenuItem_ATE_Click(object sender, RoutedEventArgs e)
        {
            ATEWindow ateWindow = new ATEWindow();
            ateWindow.Show();
        }

        private async void DoReport_Click(object sender, RoutedEventArgs e)
        {
            PopupWindow popup = new PopupWindow() { Title = "处理中", Message = "请耐心等待..." };
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
                string ReportName = text_rootReportPath.Text;
                if (!File.Exists(ReportName))
                {
                    throw new FileNotFoundException("报告文件不存在");
                }
                await ReadInfoFromOverview(ReportName);
                _logger.Info("报告概览读取完成");

                thermalshockPage.ReadReportHeader();
                thermalshockPage.SetReportResultData();
                burninPage.ReadReportHeader();
                burninPage.SetReportResultData();
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

        private void Info_Set_Click(object sender, RoutedEventArgs e)
        {
            SetInfoToWindow();
        }

        private void btn_rootReportPath_Click(object sender, RoutedEventArgs e)
        {
            FileDialog fd = new OpenFileDialog()
            {
                Filter = "Excel文件|*.xlsx;*.xls"
            };
            _ = fd.ShowDialog();
            if (fd.FileName is string fileName && fileName != "")
            {
                RootPath = Path.GetDirectoryName(fileName);
                text_rootReportPath.Text = fileName;
                string _title = Path.GetFileName(Path.GetDirectoryName(fileName));
                Title = _title.Split(' ')[0] + " " + _title.Split('_')[1] + " ORT一键报告";
            }
        }

        private void btn_finish_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string btnTag)
            {
            }
        }
    }
}

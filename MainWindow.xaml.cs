using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using NLog;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static ORT一键报告.ReportHeader;

namespace ORT一键报告
{
    public enum ReportStatus { Pass, Fail };

    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public ReportHeaderInfo thermalReportHeaderInfo = null;
        public ReportHeaderInfo burnReportHeaderInfo = null;
        public ReportHeaderInfo emiReportHeaderInfo = null;
        public ObservableCollection<ResultDetails> DetailsList = new ObservableCollection<ResultDetails>();
        public string ATEPath { get; set; }
        public string RootPath { get; set; }
        public string TempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ORTTemp");

        public MainWindow()
        {
            InitializeComponent();
            ExcelPackage.License.SetNonCommercialPersonal("Lucas");
            Closed += Window_Closed;
            _logger.Info("设置DataGrid数据源");
            t_details_data.DataGridSource = DetailsList;
            b_details_data.DataGridSource = DetailsList;
            b_details_data.InitBurnColumns();
            t_details_data.InitThermalColumns();
            t_details_data.AddRow();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            ClearTempDir(_logger);
        }

        /* ###############################  功能函数  ################################ */

        private async Task<List<object>> ReadInfoFromOverview(string ReportName)
        {
            _logger.Info("读取报告概览...");

            DateTime t_start = DateTime.Now;
            DateTime b_start = DateTime.Now;
            string SNsCount = "3";
            UUTInfoFromExcel UUTInfos;
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
                return null;
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
            return new List<object>() { t_start, b_start, SNsCount };

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

        public void ReadReportHeader()
        {
            ReportHeaderInfo fun(ExcelWorksheet ws)
            {
                // 辅助函数: 找到issue和setup图片所在的标题行
                DataCell issueTitle = FindCellByValue(ws, "Issue Photos");
                DataCell setupTitle = FindCellByValue(ws, "Test Setup");

                ReportHeaderInfo reportHeaderInfo = new ReportHeaderInfo
                {
                    TESTED_BY = FindInfoByText(ws, "TESTED BY"),
                    APPROVED_BY = FindInfoByText(ws, "APPROVED BY"),
                    PROJECT_NAME = FindInfoByText(ws, "PROJECT NAME"),
                    TEST_STAGE = FindInfoByText(ws, "TEST STAGE"),
                    TestDescription = FindInfoByText(ws, "Test Description"),
                    Test_Description_Pic = GetPicturesInRange(ws, 6, 1, 10),
                    Issue_Photos_Pics = issueTitle is null ? null : GetPicturesInRange(ws, issueTitle.Row, 1, issueTitle.Row + 10),
                    Test_Setup_Pics = setupTitle is null ? null : GetPicturesInRange(ws, setupTitle.Row, 1, setupTitle.Row + 10),
                };
                return reportHeaderInfo;
            }

            _logger.Info("读取报告表头...");
            FileInfo thermalFileInfo = new FileInfo(GetTemplatePath(RootPath, "Thermal Shock"));
            using (ExcelPackage package = new ExcelPackage(thermalFileInfo))
            {
                var ws = package.Workbook.Worksheets[0];

                thermalReportHeaderInfo = fun(ws);
                _logger.Info("thermal表头读取完成");
            }

            FileInfo burnFileInfo = new FileInfo(GetTemplatePath(RootPath, "Burn"));
            using (ExcelPackage package = new ExcelPackage(burnFileInfo))
            {
                var ws = package.Workbook.Worksheets[0];

                burnReportHeaderInfo = fun(ws);
                _logger.Info("burn表头读取完成");
            }

            FileInfo emiFileInfo = new FileInfo(GetTemplatePath(RootPath, "EMI"));
            using (ExcelPackage package = new ExcelPackage(emiFileInfo))
            {
                var ws = package.Workbook.Worksheets[0];

                emiReportHeaderInfo = fun(ws);
                _logger.Info("EMI表头读取完成");
            }
        }

        private void SetInfoToWindow()
        {
            fun(t_ReportHeader, thermalReportHeaderInfo, widget_pic_t);
            fun(b_ReportHeader, burnReportHeaderInfo, widget_pic_b);
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

        public void ExcelAddPicture(ExcelWorksheet ws, string picName, DataCell pics, string TopLeft, string rpType)
        {
            if (pics.Images.Count <= 0)
            {
                return;
            }
            var start = new ExcelAddress(TopLeft).Start;
            int startRow = start.Row;
            int startCol = start.Column;
            for (int i = 0; i < pics.Images.Count; i++)
            {
                string picPath = System.IO.Path.Combine(TempPath, picName + "_" + i + ".png");
                if (File.Exists(picPath))
                {
                    var temp = picPath.Split('.');
                    picPath = temp[0] + "_" + i + "." + temp[1];
                }
                ImageSaverLegacy.SaveImageSourceToFile(pics.Images[i].ImageSrc, picPath, "png");
                ExcelPicture test_desc_pic_excel = ws.Drawings.AddPicture(picName + "_" + i, picPath);
                test_desc_pic_excel.SetSize(300, 220);
                if (rpType == "burn")
                {
                    test_desc_pic_excel.SetPosition(startRow, 0, startCol + (i * 4), -18 + (i * 72));
                }
                else
                {
                    test_desc_pic_excel.SetPosition(startRow, 10, startCol + (i * 4), -24 + (i * 44));
                }
            }
        }

        public void ReportFinish(string ReportType, List<object> HeaderInfoList, ReportHeaderInfo reportHeaderInfo)
        {
            _logger.Info($"{ReportType}报告生成中...");
            string saveReportPath;
            try
            {
                string currentPath = Directory.GetCurrentDirectory();
                FileInfo reportFI = new FileInfo(GetTemplatePath(System.IO.Path.Combine(currentPath, "Templates"), ReportType));
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    FileName = reportFI.Name,
                    Filter = "Excel文件|*.xlsx;*.xls",
                    InitialDirectory = RootPath
                };
                ExcelPackage package = new ExcelPackage(reportFI);
                ExcelWorkbook wb = package.Workbook;
                ExcelWorksheet ws = wb.Worksheets[0];
                ExcelWorksheet ws_setup = wb.Worksheets[1];

                // 1.表头信息
                _logger.Info("处理表头");
                for (int r = 1; r <= 8; r++)
                {
                    ws.Cells[ws_setup.Cells[r, 1].Text].Value = HeaderInfoList[r - 1];
                }

                // 2.单体数据
                _logger.Info("处理单体数据");
                List<object> detailInfoList = new List<object>
                {
                    DetailsList.Select(r => r.BIroom).ToList(),
                    DetailsList.Select(r => r.BIarea).ToList(),
                    DetailsList.Select(r => r.BIplace).ToList(),
                    DetailsList.Select(r => r.SN).ToList(),
                    DetailsList.Select(r => r.WorkOrder).ToList(),
                    DetailsList.Select(r => r.Version).ToList(),
                    DetailsList.Select(r => r.DC).ToList(),
                    DetailsList.Select(r => r.InspectionPrev).ToList(),
                    DetailsList.Select(r => r.InspectionAfter).ToList(),
                    DetailsList.Select(r => r.FunPrev).ToList(),
                    DetailsList.Select(r => r.FunAfter).ToList(),
                    DetailsList.Select(r => r.HiPot).ToList(),
                };
                if (ReportType.ToLower().Contains("thermal"))
                {
                    detailInfoList.RemoveRange(0, 3);
                }

                int _detail_start_row = 13; //setup表detail的起始行
                for (int r = _detail_start_row; r < ws_setup.Dimension.End.Row; r++)
                {
                    ExcelAddress address = new ExcelAddress(ws_setup.Cells[r, 1].Text);
                    var Rp_row = address.Start.Row;
                    var Rp_col = address.Start.Column;
                    if (detailInfoList[r - _detail_start_row] is List<string> detailInfo)
                    {
                        ws.Cells[Rp_row, Rp_col].Value = detailInfo[0];
                        ws.Cells[Rp_row + 1, Rp_col].Value = detailInfo[1];
                        ws.Cells[Rp_row + 2, Rp_col].Value = detailInfo[2];
                    }
                    else if (detailInfoList[r - _detail_start_row] is List<ReportStatus> detailStatus)
                    {
                        ws.Cells[Rp_row, Rp_col].Value = detailStatus[0].ToString();
                        ws.Cells[Rp_row + 1, Rp_col].Value = detailStatus[1].ToString();
                        ws.Cells[Rp_row + 2, Rp_col].Value = detailStatus[2].ToString();
                    }
                }

                // 3.图片和OLE对象
                _logger.Info("处理图片和OLE对象");
                ExcelAddPicture(ws, "Issue_Photos", reportHeaderInfo.Issue_Photos_Pics, ws_setup.Cells["A11"].Text, "burn");
                ExcelAddPicture(ws, "Test_Setup", reportHeaderInfo.Test_Setup_Pics, ws_setup.Cells["A12"].Text, "burn");

                string ate_Addr = ws_setup.Cells["A9"].Text;
                wb.Worksheets.Delete(ws_setup); // 删除设置表

                saveReportPath = saveFileDialog.ShowDialog() == true
                    ? saveFileDialog.FileName
                    : System.IO.Path.Combine(Directory.GetCurrentDirectory(), reportFI.Name);
                saveReportPath = System.IO.Path.GetFullPath(saveReportPath);
                package.SaveAs(saveReportPath);
                EmbedOleObjectWithInterop(_logger, saveReportPath, ATEPath, ate_Addr);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{ReportType}模版生成失败");
                MessageBox.Show(ex + $"{ReportType}模版生成失败");
                return;
            }
            _logger.Info($"{ReportType}报告生成完成, 保存在{saveReportPath}");
            MessageBox.Show($"{ReportType}报告生成完成, 保存在{saveReportPath}", "成功");
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
                List<object> res = await ReadInfoFromOverview(ReportName);
                t_ReportHeader.datepicker_start.SelectedDate = (DateTime?)res[0];
                b_ReportHeader.datepicker_start.SelectedDate = (DateTime?)res[1];
                t_details_data.UUT_Count = b_details_data.UUT_Count = int.Parse((string)res[2]);

                _logger.Info("报告概览读取完成");
                ReadReportHeader();
                SetInfoToWindow();
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

        private void btn_ATEDatas_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string textName)
            {
                FileDialog fileDialog = new OpenFileDialog
                {
                    Filter = "ATE数据|*.xlsx;*.xls"
                };
                _ = fileDialog.ShowDialog();

                ATEPath = fileDialog.FileName;
                if (FindName(textName) is TextBlock textBlock)
                {
                    textBlock.Text = ATEPath;
                }
            }
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
                RootPath = System.IO.Path.GetDirectoryName(fileName);
                text_rootReportPath.Text = fileName;
            }
        }

        private async void btn_finish_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string btnTag)
            {
                button.IsEnabled = false;

                _ = t_details_data.details_data.CommitEdit();
                _ = b_details_data.details_data.CommitEdit();
                ReportHeaderWidget _ReportHeader;
                ReportHeaderInfo _ReportHeaderInfo;
                switch (btnTag)
                {
                    case "thermal":
                        _ReportHeader = t_ReportHeader;
                        _ReportHeaderInfo = thermalReportHeaderInfo;
                        break;
                    case "burn":
                        _ReportHeader = b_ReportHeader;
                        _ReportHeaderInfo = burnReportHeaderInfo;
                        break;
                    default:
                        _ReportHeader = t_ReportHeader;
                        _ReportHeaderInfo = thermalReportHeaderInfo;
                        break;
                }
                List<object> HeaderInfoList = new List<object> {
                    _ReportHeader.TestedBy,
                    _ReportHeader.ApprovedBy,
                    _ReportHeader.ProjectName,
                    _ReportHeader.TestStage,
                    _ReportHeader.datepicker_start.SelectedDate,
                    _ReportHeader.datepicker_end.SelectedDate,
                    (bool)_ReportHeader.rbtn_testPass.IsChecked ? "Pass" : "Fail",
                    _ReportHeader.text_TestDescription.Text,
                };
                PopupWindow popup = new PopupWindow() { Title = "保存报告", Message = "处理中..." };
                try
                {
                    popup.Show();
                    await Task.Run(() => { ReportFinish(btnTag, HeaderInfoList, _ReportHeaderInfo); });
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "保存报告时出现错误");
                    popup.Message = "保存报告时出现错误" + ex.Message;
                }
                finally
                {
                    popup.Close();
                    button.IsEnabled = true;
                }
            }
        }
    }
}

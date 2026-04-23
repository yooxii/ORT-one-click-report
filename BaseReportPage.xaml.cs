using System;
using System.Collections.Generic;
using System.Linq;
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
using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using NLog;
using System.Collections.ObjectModel;
using System.IO;

namespace ORT一键报告
{
    /// <summary>
    /// BaseReportPage.xaml 的交互逻辑
    /// </summary>
    public partial class BaseReportPage : UserControl
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public ReportHeaderInfo reportHeaderInfo = null;
        public ObservableCollection<ResultDetails> DetailsList = new ObservableCollection<ResultDetails>();
        public string ATEPath { get; set; }
        public string RootPath { get; set; }
        public string RootReportPath { get; set; }
        public string TempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ORTTemp");

        public BaseReportPage()
        {
            InitializeComponent();
            _logger.Info("设置DataGrid数据源");
            details_data.DataGridSource = DetailsList;
            details_data.InitThermalColumns();
            details_data.AddRow();
        }

        /* ###############################  功能函数  ################################ */

        private void SetInfoToWindow()
        {
            fun(ReportHeader, reportHeaderInfo, widget_pic);

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


        private async void btn_finish_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string btnTag)
            {
                button.IsEnabled = false;

                _ = details_data.details_data.CommitEdit();
                List<object> HeaderInfoList = new List<object> {
                    ReportHeader.TestedBy,
                    ReportHeader.ApprovedBy,
                    ReportHeader.ProjectName,
                    ReportHeader.TestStage,
                    ReportHeader.datepicker_start.SelectedDate,
                    ReportHeader.datepicker_end.SelectedDate,
                    (bool)ReportHeader.rbtn_testPass.IsChecked ? "Pass" : "Fail",
                    ReportHeader.text_TestDescription.Text,
                };
                PopupWindow popup = new PopupWindow() { Title = "保存报告", Message = "处理中..." };
                try
                {
                    popup.Show();
                    await Task.Run(() => { ReportFinish(btnTag, HeaderInfoList, reportHeaderInfo); });
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

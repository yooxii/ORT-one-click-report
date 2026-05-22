using NLog;
using OfficeOpenXml;
using ORT一键报告.ViewModels;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using static ORT一键报告.ReportUtils;

namespace ORT一键报告
{
    /// <summary>
    /// BaseReportPage.xaml 的交互逻辑
    /// </summary>
    public partial class BaseReportPage : UserControl
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public BaseReportPageViewModel BaseReportPageVM { get; }
        public ReportHeaderViewModel ReportHeaderInfo { get; set; }


        public string ReportType
        {
            get => (string)GetValue(ReportTypeProperty);
            set => SetValue(ReportTypeProperty, value);
        }

        public static readonly DependencyProperty ReportTypeProperty =
            DependencyProperty.Register("ReportType", typeof(string), typeof(BaseReportPage), new PropertyMetadata("thermal"));


        public int TestTime
        {
            get => (int)GetValue(TestTimeProperty);
            set => SetValue(TestTimeProperty, value);
        }

        public static readonly DependencyProperty TestTimeProperty =
            DependencyProperty.Register("TestTime", typeof(int), typeof(BaseReportPage), new PropertyMetadata(1));



        public BaseReportPage()
        {
            InitializeComponent();
            Service service = new();
            BaseReportPageVM = new(service)
            {
                ReportType = ReportType
            };
            ReportHeaderInfo = BaseReportPageVM.ReportHeaderVM;
            DataContext = BaseReportPageVM;
        }

        public void InitReportPage()
        {
            _logger.Info($"设置{ReportType}-DataGrid的数据源");
            details_data.InitColumns(ReportType);
            details_data.AddRow();
        }

        /* ###############################  功能函数  ################################ */

        public void SetReportResultData()
        {
            if (BaseReportPageVM.DetailsList == null)
            {
                BaseReportPageVM.DetailsList = new ObservableCollection<ResultDetails>();
            }
            BaseReportPageVM.DetailsList.Clear();
            UUTInfoFromExcel _UUTInfos = MainWindow.UUTInfos;
            foreach (TestItemInfo testItem in _UUTInfos.TestItems)
            {
                if (testItem.TestItemName.ToLower().Contains(ReportType.ToLower()))
                {
                    ReportHeader.datepicker_start.SelectedDate = DateTime.Parse(testItem.Date);
                    ReportHeaderInfo.TestStart = DateTime.Parse(testItem.Date);
                    ReportHeaderInfo.TestEnd = DateTime.Parse(testItem.Date).AddDays(TestTime);
                }
            }
            foreach (string uutSNs in _UUTInfos.SNs)
            {
                BaseReportPageVM.DetailsList.Add(new ResultDetails()
                {
                    BIroom = "1F Chamber",
                    SN = uutSNs,
                    WorkOrder = _UUTInfos.WorkerNo,
                    Version = _UUTInfos.Revision,
                    DC = _UUTInfos.DC,
                    InspectionPrev = ReportStatus.Pass,
                    FunPrev = ReportStatus.Pass,
                    InspectionAfter = ReportStatus.Pass,
                    FunAfter = ReportStatus.Pass,
                    HiPot = ReportStatus.Pass,
                    Comments = ""
                });
            }
        }

        public void ReadReportHeader()
        {
            _logger.Info($"读取{ReportType}报告表头...");
            FileInfo thermalFileInfo = new(GetTemplatePath(MainWindow.RootPath, ReportType));
            using (ExcelPackage package = new(thermalFileInfo))
            {
                ExcelWorksheet ws = package.Workbook.Worksheets[0];

                ReadReportHeaderInfo(ws, ReportHeaderInfo);
                _logger.Info($"{ReportType}表头读取完成");
            }
            SetInfoToWindow();
        }

        private void SetInfoToWindow()
        {
            static void SetPics(List<ExcelPictureInfo> _pics, List<Image> images)
            {
                for (int i = 0; i < _pics.Count && i < 3; i++)
                {
                    images[i].Source = _pics[i].ImageSrc;
                }
            }

            ReportHeader.ApprovedBy = ReportHeaderInfo.APPROVED_BY.Data;
            ReportHeader.TestedBy = ReportHeaderInfo.TESTED_BY.Data;
            ReportHeader.ProjectName = ReportHeaderInfo.PROJECT_NAME.Data;
            ReportHeader.TestStage = ReportHeaderInfo.TEST_STAGE.Data;
            ReportHeader.TextTestDescription = ReportHeaderInfo.TestDescription.Data;

            if (ReportHeaderInfo.Issue_Photos_Pics != null)
            {
                SetPics(ReportHeaderInfo.Issue_Photos_Pics.Images, new List<Image> { widget_pic.issue_image1, widget_pic.issue_image2, widget_pic.issue_image3 });
            }
            if (ReportHeaderInfo.Test_Setup_Pics != null)
            {
                SetPics(ReportHeaderInfo.Test_Setup_Pics.Images, new List<Image> { widget_pic.setup_image1, widget_pic.setup_image2, widget_pic.setup_image3 });
            }
        }

        /* ###############################  事件函数  ################################ */

        private void Info_Set_Click(object sender, RoutedEventArgs e)
        {
            SetInfoToWindow();
        }

        //private void btn_ATEDatas_Click(object sender, RoutedEventArgs e)
        //{
        //    if (sender is Button button && button.Tag is string textName)
        //    {
        //        FileDialog fileDialog = new OpenFileDialog
        //        {
        //            Filter = "ATE数据|*.xlsx;*.xls"
        //        };
        //        _ = fileDialog.ShowDialog();

        //        ATEPath = fileDialog.FileName;
        //        if (FindName(textName) is TextBlock textBlock)
        //        {
        //            textBlock.Text = ATEPath;
        //        }
        //    }
        //}

        //private async void btn_finish_Click(object sender, RoutedEventArgs e)
        //{
        //    if (sender is Button button && button.Tag is string btnTag)
        //    {
        //        button.IsEnabled = false;

        //        _ = details_data.details_data.CommitEdit();
        //        List<object> HeaderInfoList = new()
        //        {
        //            ReportHeader.TestedBy,
        //            ReportHeader.ApprovedBy,
        //            ReportHeader.ProjectName,
        //            ReportHeader.TestStage,
        //            ReportHeader.datepicker_start.SelectedDate,
        //            ReportHeader.datepicker_end.SelectedDate,
        //            (bool)ReportHeader.rbtn_testPass.IsChecked ? "Pass" : "Fail",
        //            ReportHeader.text_TestDescription.Text,
        //        };
        //        PopupWindow popup = new() { Title = "保存报告", Message = "处理中..." };
        //        try
        //        {
        //            popup.Show();
        //            await Task.Run(() => { ReportFinish(btnTag, HeaderInfoList, reportHeaderInfo); });
        //        }
        //        catch (Exception ex)
        //        {
        //            _logger.Error(ex, "保存报告时出现错误");
        //            popup.Message = "保存报告时出现错误" + ex.Message;
        //        }
        //        finally
        //        {
        //            popup.Close();
        //            button.IsEnabled = true;
        //        }
        //    }
        //}
    }
}

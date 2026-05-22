using NLog;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using static ORT一键报告.ReportUtils;

namespace ORT一键报告
{

    /// <summary>
    /// EMIReportPage.xaml 的交互逻辑
    /// </summary>
    public partial class EMIReportPage : UserControl
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public static ViewModels.EMIViewModel emiVM;

        public ViewModels.ReportHeaderViewModel ReportHeaderInfo { get; set; }

        public string ReportType
        {
            get => (string)GetValue(ReportTypeProperty);
            set => SetValue(ReportTypeProperty, value);
        }

        public static readonly DependencyProperty ReportTypeProperty =
            DependencyProperty.Register("ReportType", typeof(string), typeof(EMIReportPage), new PropertyMetadata("EMI"));

        public int TestTime
        {
            get => (int)GetValue(TestTimeProperty);
            set => SetValue(TestTimeProperty, value);
        }

        public static readonly DependencyProperty TestTimeProperty =
            DependencyProperty.Register("TestTime", typeof(int), typeof(EMIReportPage), new PropertyMetadata(1));

        public EMIReportPage()
        {
            InitializeComponent();
            Service service = new();
            emiVM = new(service);
            ReportHeaderInfo = emiVM.ReportHeaderVM;
            DataContext = emiVM;
        }

        /* ###############################  功能函数  ################################ */

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
    }
}

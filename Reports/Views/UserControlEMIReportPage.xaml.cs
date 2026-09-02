using Microsoft.Extensions.DependencyInjection;
using NLog;
using OfficeOpenXml;
using ORT一键报告.Models;
using ORT一键报告.Reports.ViewModels;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using static ORT一键报告.Utils.Report;

namespace ORT一键报告.Reports.Views
{

    /// <summary>
    /// EMIReportPage.xaml 的交互逻辑
    /// </summary>
    public partial class EMIReportPage : UserControl
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public static EMIReportViewModel emiVM;

        public ReportHeaderViewModel ReportHeaderInfo { get; set; }

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
            emiVM = App.ServiceProvider.GetRequiredService<EMIReportViewModel>();
            ReportHeaderInfo = emiVM.ReportHeaderVM;
            ReportHeader.DataContext = ReportHeaderInfo;
            DataContext = emiVM;
        }

        /* ###############################  功能函数  ################################ */

        public void ReadReportHeader()
        {
            _logger.Info($"读取{ReportType}报告表头...");
            ReportService reportService = App.ServiceProvider.GetRequiredService<ReportService>();
            string templatePath = GetTemplatePath(reportService.RootPath, ReportType);
            if (string.IsNullOrWhiteSpace(templatePath) || !File.Exists(templatePath))
            {
                _logger.Warn($"未找到{ReportType}报告模板，跳过读取该报告表头");
                return;
            }
            FileInfo fileInfo = new(templatePath);
            using (ExcelPackage package = new(fileInfo))
            {
                ExcelWorksheet ws = package.Workbook.Worksheets[0];

                ReadReportHeaderInfo(ws, ReportHeaderInfo);
                _logger.Info($"{ReportType}表头读取完成");
            }
            UUTInfoFromExcel _UUTInfos = reportService.UUTInfos;
            if (_UUTInfos == null)
            {
                _logger.Warn($"{ReportType}报告：UUTInfos 为空，跳过 EMI 数据填充");
                return;
            }
            emiVM.DC = _UUTInfos.DC;
            emiVM.Version = _UUTInfos.Revision;
            emiVM.WorkOrder = _UUTInfos.WorkOrder;
            foreach (TestItemInfo testItem in _UUTInfos.TestItems ?? [])
            {
                if (testItem.TestItemName?.ToLower().Contains(ReportType.ToLower()) == true)
                {
                    if (!DateTime.TryParse(testItem.Date, out DateTime parsedDate))
                    {
                        _logger.Warn($"{ReportType}报告：测试项目 {testItem.TestItemName} 的日期无效（{testItem.Date}），跳过日期填充");
                        continue;
                    }
                    ReportHeader.datepicker_start.SelectedDate = parsedDate;
                    ReportHeaderInfo.TestStart = parsedDate;
                    ReportHeaderInfo.TestEnd = parsedDate.AddDays(TestTime);
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

            ReportHeader.ApprovedBy = ReportHeaderInfo.APPROVED_BY?.Data ?? "";
            ReportHeader.TestedBy = ReportHeaderInfo.TESTED_BY?.Data ?? "";
            ReportHeader.ProjectName = ReportHeaderInfo.PROJECT_NAME?.Data ?? "";
            ReportHeader.TestStage = ReportHeaderInfo.TEST_STAGE?.Data ?? "";
            ReportHeader.TextTestDescription = ReportHeaderInfo.TestDescription?.Data ?? "";

            if (ReportHeaderInfo.Issue_Photos_Pics != null)
            {
                SetPics(ReportHeaderInfo.Issue_Photos_Pics.Images, [widget_pic.issue_image1, widget_pic.issue_image2, widget_pic.issue_image3]);
            }
            if (ReportHeaderInfo.Test_Setup_Pics != null)
            {
                SetPics(ReportHeaderInfo.Test_Setup_Pics.Images, [widget_pic.setup_image1, widget_pic.setup_image2, widget_pic.setup_image3]);
            }
        }

        private Window GetRootWindow(FrameworkElement framework)
        {
            if (framework is Window fw)
            {
                return fw;
            }
            else if (framework.Parent is FrameworkElement fe)
            {
                return GetRootWindow(fe);
            }
            else
            {
                return null;
            }
        }

        private void BTNEMISetup_Click(object sender, RoutedEventArgs e)
        {
            EMIReportSetup emisetup = new()
            {
                DataContext = emiVM.EMISetupVM,
                Owner = GetRootWindow(this)
            };
            emiVM.EMISetupVM.TemplatePath = emiVM.TemplatePath;
            emiVM.EMISetupVM.LoadFromExcel();
            emisetup.Show();
        }
    }
}

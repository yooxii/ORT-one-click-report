using Microsoft.Extensions.DependencyInjection;
using NLog;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using ExcelNpoi = ORT一键报告.Utils.ExcelNpoi;
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

        /// <summary>「数据路径」标签连击计数（连击三下打开 EMI 数据处理对话框）</summary>
        private int _dataPathClickCount;
        private DateTime _lastDataPathClick = DateTime.MinValue;

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

        /// <summary>
        /// 「数据路径」标签连击三下打开 EMI 数据处理（原 Alert 按钮入口已隐藏）。
        /// 时间窗口 800ms 内累计三次才算连击，避免零散点击误触发。
        /// </summary>
        private void Label_DataPath_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if ((DateTime.Now - _lastDataPathClick).TotalMilliseconds > 800)
            {
                _dataPathClickCount = 0;
            }
            _lastDataPathClick = DateTime.Now;
            _dataPathClickCount++;
            if (_dataPathClickCount < 3)
            {
                return;
            }
            _dataPathClickCount = 0;
            _logger.Info("检测到数据路径标签连击三下，打开 EMI 数据处理对话框");
            if (emiVM.AlertTimeCommand.CanExecute(null))
            {
                emiVM.AlertTimeCommand.Execute(null);
            }
        }

        /* ###############################  功能函数  ################################ */

        public void ReadReportHeader()
        {
            _logger.Info($"读取{ReportType}报告表头...");
            ReportService reportService = App.ServiceProvider.GetRequiredService<ReportService>();
            // 测试信息与测试图片以"该计划绑定的报告文件夹"里的本地报告文件为准，没有才回退模板
            string reportFilePath = GetTemplatePath(reportService.MatchedReportDir, ReportType);
            bool fromReportFile = !string.IsNullOrWhiteSpace(reportFilePath) && File.Exists(reportFilePath);
            // 从计划表进入时，报告文件夹里没有这类报告 → 表头与图片一律置空（不再回退模板，避免带出旧数据）
            if (reportService.EnteredFromPlan && !fromReportFile)
            {
                _logger.Warn($"{ReportType}：绑定的报告文件夹里没有该类报告，测试信息与图片置空（{reportService.MatchedReportDir}）");
                ResetReportHeaderInfo(ReportHeaderInfo);
                SetInfoToWindow();
                return;
            }
            string templatePath = fromReportFile ? reportFilePath : GetTemplatePath(reportService.RootPath, ReportType);
            if (string.IsNullOrWhiteSpace(templatePath) || !File.Exists(templatePath))
            {
                _logger.Warn($"未找到{ReportType}报告模板，跳过读取该报告表头");
                return;
            }
            XSSFWorkbook wb = ExcelNpoi.OpenRead(templatePath);
            try
            {
                ISheet ws = ExcelNpoi.SheetAt(wb, 0);

                ReadReportHeaderInfo(ws, ReportHeaderInfo);
                _logger.Info(fromReportFile
                    ? $"{ReportType}表头读取自本地报告文件：{templatePath}"
                    : $"{ReportType}表头读取自模板：{templatePath}");
            }
            finally
            {
                wb.Close();
            }
            // 报告文件里的"TEST PERIOD"优先；其次用报告概览文件夹名里的周号（WK####）；最后才用测试项目日期
            DateTime? periodFromReport = ReportHeaderInfo.TestStart;
            if (periodFromReport == null && reportService.UUTInfos?.TestStart != null)
            {
                periodFromReport = reportService.UUTInfos.TestStart;
                _logger.Info($"{ReportType}测试周期取自报告概览周号 {reportService.UUTInfos.TestPeriod}：{periodFromReport:yyyy-MM-dd}");
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
                    if (periodFromReport == null)
                    {
                        ReportHeader.datepicker_start.SelectedDate = parsedDate;
                        ReportHeaderInfo.TestStart = parsedDate;
                        ReportHeaderInfo.TestEnd = parsedDate.AddDays(TestTime);
                    }
                }
            }
            if (periodFromReport != null)
            {
                ReportHeader.datepicker_start.SelectedDate = periodFromReport;
                ReportHeaderInfo.TestStart = periodFromReport;
                ReportHeaderInfo.TestEnd = periodFromReport.Value.AddDays(TestTime);
            }
            SetInfoToWindow();
        }

        private void SetInfoToWindow()
        {
            static void SetPics(List<ExcelPictureInfo> _pics, List<Image> images)
            {
                // 图片不足 3 张时也要把多余槽位清空，否则界面会残留上一份报告的图片
                List<ExcelPictureInfo> pics = _pics ?? [];
                for (int i = 0; i < images.Count && i < 3; i++)
                {
                    images[i].Source = i < pics.Count ? pics[i].ImageSrc : null;
                }
            }

            ReportHeader.ApprovedBy = ReportHeaderInfo.APPROVED_BY?.Data ?? "";
            ReportHeader.TestedBy = ReportHeaderInfo.TESTED_BY?.Data ?? "";
            ReportHeader.ProjectName = ReportHeaderInfo.PROJECT_NAME?.Data ?? "";
            ReportHeader.TestStage = ReportHeaderInfo.TEST_STAGE?.Data ?? "";
            ReportHeader.TextTestDescription = ReportHeaderInfo.TestDescription?.Data ?? "";

            SetPics(ReportHeaderInfo.Issue_Photos_Pics?.Images, [widget_pic.issue_image1, widget_pic.issue_image2, widget_pic.issue_image3]);
            SetPics(ReportHeaderInfo.Test_Setup_Pics?.Images, [widget_pic.setup_image1, widget_pic.setup_image2, widget_pic.setup_image3]);
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

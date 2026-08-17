using Microsoft.Extensions.DependencyInjection;
using NLog;
using OfficeOpenXml;
using ORT一键报告.Main.Views;
using ORT一键报告.Models;
using ORT一键报告.Reports.ViewModels;
using ORT一键报告.Services;
using ORT一键报告.ViewModels;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using static ORT一键报告.Utils.Report;

namespace ORT一键报告.Reports.Views
{
    public enum ReportStatus { Pass, Fail };

    /// <summary>
    /// WindowMainReport.xaml 的交互逻辑
    /// </summary>
    public partial class WindowMainReport : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public ReportService ReportService { get; }
        public MainReportViewModel MainVM { get; set; }

        public static SettingsViewModel SettingsVM { get; set; } = new();

        private readonly Dictionary<string, object> defaultSetup = new() {
            {"路径对话框初始目录", new Dictionary<string, object> {
                {"BI EMI 报告","\\\\bnt56\\品保部\\ORT實驗資料\\13. 臨時試驗報告\\BI EMI"},
                {"BI ATE Data", "\\\\bnt56\\品保部\\ORT實驗資料\\13. 臨時試驗報告\\BI ATE Data" },
                {"BI Picture","\\\\bnt56\\品保部\\ORT實驗資料\\13. 臨時試驗報告\\BI Picture" }
            } },
        };

        public WindowMainReport()
        {
            InitializeComponent();

            ReportService = App.ServiceProvider.GetRequiredService<ReportService>();
            MainVM = App.ServiceProvider.GetRequiredService<MainReportViewModel>();
            DataContext = MainVM;

            Loaded += ReportHeader_Loaded;

            ReportService.TemplateDir = Path.Combine(Directory.GetCurrentDirectory(), "Templates");
            ReportService.TempPath = Path.Combine(Path.GetTempPath(), "ORTTemp");
        }

        private void ReportHeader_Loaded(object sender, RoutedEventArgs e)
        {
            thermalshockPage.InitReportPage();
            burninPage.InitReportPage();
        }


        /* ###############################  事件函数  ################################ */
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
                await MainVM.ReadInfoFromOverview(ReportName);
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

        private void MenuItem_ATE_Click(object sender, RoutedEventArgs e)
        {
            ATEWindow ateWindow = new()
            {
                Owner = this
            };
            ateWindow.Show();
        }

        private void MenuItem_ReportTemplate_Click(object sender, RoutedEventArgs e)
        {
            WindowReportTemplate windowReportTemplate = new()
            {
                Owner = this
            };
            windowReportTemplate.Show();
        }

        private void MenuItem_ReturnLine_Click(object sender, RoutedEventArgs e)
        {
            WindowReturnLine windowReturnLine = new();
            windowReturnLine.Show();
        }

        private void MenuItem_ViewLog_Click(object sender, RoutedEventArgs e)
        {
            WindowLog windowLog = new()
            {
                Owner = this
            };
            windowLog.Show();
        }
    }
}

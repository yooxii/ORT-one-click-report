using NLog;
using OfficeOpenXml;
using ORT一键报告.Models;
using ORT一键报告.Reports.Views;
using ORT一键报告.ViewModels;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using static ORT一键报告.Utils.Report;

namespace ORT一键报告
{

    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public MainViewModel MainVM { get; set; }

        public static SettingsViewModel SettingsVM { get; set; } = new();

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
            MainVM = new();
            DataContext = MainVM;

            Closed += Window_Closed;
        }


        private void Window_Closed(object sender, EventArgs e)
        {
            ClearTempDir();
        }

        /* ###############################  事件函数  ################################ */


        private void MenuItem_MainSetup_Click(object sender, RoutedEventArgs e)
        {
            MainSettingsWindow mainSettingsWindow = new()
            {
                Owner = this
            };
            mainSettingsWindow.Show();
        }

        private void MenuItem_Quit_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Button_YiJianBaoGao_Click(object sender, RoutedEventArgs e)
        {
            WindowMainReport windowMainReport = new();
            windowMainReport.Show();
        }

        private void Button_Plans_Click(object sender, RoutedEventArgs e)
        {
            
        }
    }
}
using ORT一键报告.Plans.ViewModels;
using ORT一键报告.Services;
using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Documents;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// WindowFlow.xaml 的交互逻辑：流程查看窗口（只读展示，不在此界面推进流程）。
    /// 从领退表/计划表的右键菜单打开，步骤状态由 FlowViewModel 从现有登记数据推导。
    /// </summary>
    public partial class WindowFlow : Window
    {
        public WindowFlow(FlowViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }

        /// <summary>
        /// 步骤里的可点击链接（如「打開序列號文件」）：用系统默认程序打开附件
        /// </summary>
        private void FlowLink_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not Hyperlink link || link.Tag is not string path)
            {
                return;
            }
            if (!File.Exists(path))
            {
                _ = MessageBox.Show(string.Format(LocalizationHelper.Get("Msg_FileNotFoundFormat"), path),
                    LanguageService.Get("Cap_Info"));
                return;
            }
            try
            {
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                _ = MessageBox.Show($"打开文件失败:\n{ex.Message}", LanguageService.Get("Cap_Error"));
            }
        }

        private void Btn_Close_Click(object sender, RoutedEventArgs e) => Close();
    }
}

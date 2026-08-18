using Microsoft.Extensions.DependencyInjection;
using Newtonsoft.Json;
using NLog;
using ORT一键报告.Plans.ViewModels;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// WindowPlans.xaml 的交互逻辑
    /// </summary>
    public partial class WindowPlans : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public PlansViewModel PlansVM { get; }

        /// <summary>
        /// 列顺序布局文件（Data目录下），窗口关闭时保存，下次开启恢复
        /// </summary>
        private string LayoutFilePath => Path.Combine(
            App.ServiceProvider.GetService(typeof(DatabaseService)) is DatabaseService db ? db.DataDir : "", "plans_layout.json");

        public WindowPlans()
        {
            InitializeComponent();
            PlansVM = App.ServiceProvider.GetRequiredService<PlansViewModel>();
            DataContext = PlansVM;

            Loaded += (s, e) => RestoreColumnLayout();
            Closing += (s, e) => SaveColumnLayout();
        }

        /// <summary>
        /// 打开回线转移单工具（从一键报告迁移至此）
        /// </summary>
        private void Btn_ReturnLine_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            WindowReturnLine windowReturnLine = new()
            {
                Owner = this
            };
            windowReturnLine.Show();
        }

        /* ###############################  列顺序持久化  ################################ */

        /// <summary>
        /// 保存当前列显示顺序（列头名 -> DisplayIndex）
        /// </summary>
        private void SaveColumnLayout()
        {
            try
            {
                List<KeyValuePair<string, int>> layout = dg_plans.Columns
                    .OrderBy(c => c.DisplayIndex)
                    .Select(c => new KeyValuePair<string, int>(c.Header?.ToString() ?? "", c.DisplayIndex))
                    .ToList();
                File.WriteAllText(LayoutFilePath, JsonConvert.SerializeObject(layout, Formatting.Indented));
            }
            catch (Exception ex)
            {
                _logger.Warn($"保存列顺序布局失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 恢复上次保存的列显示顺序；文件不存在或不匹配时保持默认
        /// </summary>
        private void RestoreColumnLayout()
        {
            try
            {
                if (!File.Exists(LayoutFilePath))
                {
                    return;
                }
                var layout = JsonConvert.DeserializeObject<List<KeyValuePair<string, int>>>(File.ReadAllText(LayoutFilePath));
                if (layout == null)
                {
                    return;
                }
                foreach (KeyValuePair<string, int> kv in layout.OrderBy(kv => kv.Value))
                {
                    DataGridColumn col = dg_plans.Columns.FirstOrDefault(c => c.Header?.ToString() == kv.Key);
                    if (col != null && kv.Value >= 0 && kv.Value < dg_plans.Columns.Count)
                    {
                        col.DisplayIndex = kv.Value;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"恢复列顺序布局失败: {ex.Message}");
            }
        }
    }
}

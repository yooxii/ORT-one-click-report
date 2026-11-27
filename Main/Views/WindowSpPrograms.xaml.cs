using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Models;
using ORT一键报告.Plans.Views;
using ORT一键报告.Services;
using ORT一键报告.Services.Sim44;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace ORT一键报告.Main.Views
{
    /// <summary>
    /// SP程序集：把公司 MES 的常用程式集中到一个窗口里，点一下就能打开。
    /// 三类条目：
    /// 1) 本程序自己的窗口（回线工令登记 SIM44）—— 直接开 WPF 窗口；
    /// 2) 外部程式（OQC 测试记录查询 C_SII51）—— 启动它自己的 exe；
    /// 3) 占位条目（成品库存、小工令、准备领用）—— 只显示「待接入」，先不打开。
    /// 程式路径默认写在 <see cref="SpProgramCatalog"/>，厂区不同可在本机设置里覆盖。
    /// </summary>
    public partial class WindowSpPrograms : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        /// <summary>界面行：把静态清单翻译成当前语言 + 解析出实际路径与状态</summary>
        private sealed class ProgramRow
        {
            public SpProgramInfo Info { get; set; }
            public string Name { get; set; }
            public string Description { get; set; }
            public string StateText { get; set; }
            public string Path { get; set; }
            public bool IsAvailable { get; set; }
        }

        private readonly List<ProgramRow> _rows = [];

        public WindowSpPrograms()
        {
            InitializeComponent();
            Loaded += (_, _) => LoadPrograms();
        }

        private void LoadPrograms()
        {
            _rows.Clear();
            foreach (SpProgramInfo info in SpProgramCatalog.All)
            {
                string path = ResolvePath(info);
                bool available = info.Kind != SpProgramKind.Placeholder;
                _rows.Add(new ProgramRow
                {
                    Info = info,
                    Name = LanguageService.Get(info.NameKey),
                    Description = LanguageService.Get(info.DescriptionKey),
                    StateText = LanguageService.Get(available ? "SpPrograms_State_Available" : "SpPrograms_State_Todo"),
                    Path = path,
                    IsAvailable = available,
                });
            }
            dg_programs.ItemsSource = null;
            dg_programs.ItemsSource = _rows;
            dg_programs.SelectedIndex = 0;
            UpdateState();
        }

        /// <summary>取实际路径：本机设置里的覆盖优先，否则用默认值</summary>
        private static string ResolvePath(SpProgramInfo info)
        {
            try
            {
                Dictionary<string, string> overrides = LocalSettingsStore.Read()?.SpPrograms;
                if (overrides != null && overrides.TryGetValue(info.Id, out string custom)
                    && !string.IsNullOrWhiteSpace(custom))
                {
                    return custom.Trim();
                }
            }
            catch
            {
                // 设置读不到就用默认路径
            }
            return info.DefaultPath;
        }

        private void UpdateState()
        {
            if (dg_programs.SelectedItem is ProgramRow row)
            {
                txt_state.Text = row.Info.Kind == SpProgramKind.Placeholder
                    ? string.Format(LanguageService.Get("SpPrograms_Msg_TodoFormat"), row.Name)
                    : row.Path;
                btn_open.Content = LanguageService.Get("SpPrograms_Open");
            }
            else
            {
                txt_state.Text = "";
            }
            btn_open.IsEnabled = dg_programs.SelectedItem != null;
        }

        private void Dg_Programs_SelectionChanged(object sender, SelectionChangedEventArgs e) => UpdateState();

        private void Dg_Programs_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (dg_programs.SelectedItem is ProgramRow row)
            {
                Open(row);
            }
        }

        private void Btn_Open_Click(object sender, RoutedEventArgs e)
        {
            if (dg_programs.SelectedItem is ProgramRow row)
            {
                Open(row);
            }
        }

        private void Open(ProgramRow row)
        {
            switch (row.Info.Kind)
            {
                case SpProgramKind.Internal:
                    OpenInternal(row);
                    return;
                case SpProgramKind.External:
                    OpenExternal(row);
                    return;
                default:
                    MessageBox.Show(this, string.Format(LanguageService.Get("SpPrograms_Msg_TodoFormat"), row.Name),
                        LanguageService.Get("Cap_Info"), MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
            }
        }

        private void OpenInternal(ProgramRow row)
        {
            if (row.Info.Id == SpProgramCatalog.Sim44Id)
            {
                Sim44Service service = App.ServiceProvider.GetRequiredService<Sim44Service>();
                WindowSim44Return window = new(service) { Owner = this };
                window.ShowDialog();
                return;
            }
            if (row.Info.Id == SpProgramCatalog.OqcId)
            {
                WindowOqcQuery window = new() { Owner = this };
                window.ShowDialog();
                return;
            }
            MessageBox.Show(this, string.Format(LanguageService.Get("SpPrograms_Msg_TodoFormat"), row.Name),
                LanguageService.Get("Cap_Info"), MessageBoxButton.OK, MessageBoxImage.Information);
        }

        /// <summary>启动外部程式（它自己的窗口会弹出来）</summary>
        private void OpenExternal(ProgramRow row)
        {
            string path = row.Path;
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                MessageBox.Show(this, string.Format(LanguageService.Get("SpPrograms_Msg_NotFoundFormat"), path),
                    LanguageService.Get("Cap_Error"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            try
            {
                ProcessStartInfo psi = new()
                {
                    FileName = path,
                    WorkingDirectory = Path.GetDirectoryName(path) ?? string.Empty,
                    UseShellExecute = true,
                };
                using Process process = Process.Start(psi);
                _logger.Info($"SP程序集：已启动 {row.Name}（{path}）");
                // 启动成功就收起本窗口，让用户的注意力回到程式本身
                Close();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"SP程序集：启动 {path} 失败");
                MessageBox.Show(this, string.Format(LanguageService.Get("SpPrograms_Msg_LaunchFailedFormat"), ex.Message),
                    LanguageService.Get("Cap_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Btn_Close_Click(object sender, RoutedEventArgs e) => Close();
    }
}

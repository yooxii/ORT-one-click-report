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
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;

namespace ORT一键报告.Main.Views
{
    /// <summary>
    /// SP程序集：把公司 MES 的常用旧程式集中到一个窗口里，点一下就能打开。
    /// 三类条目：
    /// 1) 本程序自己的窗口（回线工令登记、OQC 测试记录查询）—— 直接开 WPF 窗口；
    /// 2) 外部程式（还没接进来的旧程式）—— 启动它自己的 exe；
    /// 3) 占位条目 —— 只显示「待接入」。
    /// 窗口是**非模态**的：主窗口随时可以点回来（抢焦点），不会被这个窗口挡住；
    /// 已经开着的旧程式，再点一次会把它的窗口调到最前面，而不是再开一份。
    /// </summary>
    public partial class WindowSpPrograms : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        /// <summary>单例：主窗口反复点「SP程序集」不会开出一堆窗口</summary>
        private static WindowSpPrograms _instance;

        /// <summary>已启动的外部程式（条目 id → 进程）</summary>
        private static readonly Dictionary<string, Process> Running = new(StringComparer.OrdinalIgnoreCase);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private const int SwRestore = 9;

        /// <summary>界面行：把静态清单翻译成当前语言 + 解析出实际路径与状态</summary>
        private sealed class ProgramRow
        {
            public SpProgramInfo Info { get; set; }
            public string Name { get; set; }
            public string Description { get; set; }
            public string StateText { get; set; }
            public string Path { get; set; }
            public bool IsAvailable { get; set; }
            public bool IsRunning { get; set; }
        }

        private readonly List<ProgramRow> _rows = [];

        public WindowSpPrograms()
        {
            InitializeComponent();
            Loaded += (_, _) => LoadPrograms();
            Activated += (_, _) => RefreshRunningState();
            Closed += (_, _) =>
            {
                if (ReferenceEquals(_instance, this))
                {
                    _instance = null;
                }
            };
        }

        /// <summary>打开（或聚焦）唯一的 SP程序集 窗口，不阻塞主窗口</summary>
        public static void ShowSingle(Window owner)
        {
            if (_instance != null)
            {
                if (_instance.WindowState == WindowState.Minimized)
                {
                    _instance.WindowState = WindowState.Normal;
                }
                _instance.Activate();
                SetForegroundWindow(new System.Windows.Interop.WindowInteropHelper(_instance).Handle);
                return;
            }
            _instance = new WindowSpPrograms { Owner = owner };
            _instance.Show();
        }

        private void LoadPrograms()
        {
            int selected = dg_programs.SelectedIndex;
            _rows.Clear();
            foreach (SpProgramInfo info in SpProgramCatalog.All)
            {
                string path = ResolvePath(info);
                bool available = info.Kind != SpProgramKind.Placeholder;
                bool running = IsRunning(info.Id);
                _rows.Add(new ProgramRow
                {
                    Info = info,
                    Name = LanguageService.Get(info.NameKey),
                    Description = LanguageService.Get(info.DescriptionKey),
                    StateText = !available
                        ? LanguageService.Get("SpPrograms_State_Todo")
                        : running
                            ? LanguageService.Get("SpPrograms_State_Running")
                            : LanguageService.Get("SpPrograms_State_Available"),
                    Path = path,
                    IsAvailable = available,
                    IsRunning = running,
                });
            }
            dg_programs.ItemsSource = null;
            dg_programs.ItemsSource = _rows;
            dg_programs.SelectedIndex = selected >= 0 && selected < _rows.Count ? selected : 0;
            UpdateState();
        }

        /// <summary>外部程式的进程如果已经退出，把记账清掉（窗口激活时刷新一次）</summary>
        private static bool IsRunning(string id)
        {
            if (!Running.TryGetValue(id, out Process process))
            {
                return false;
            }
            try
            {
                if (!process.HasExited)
                {
                    return true;
                }
            }
            catch
            {
                // 进程对象失效，按已退出处理
            }
            Running.Remove(id);
            return false;
        }

        private void RefreshRunningState()
        {
            bool changed = false;
            foreach (ProgramRow row in _rows)
            {
                bool running = row.Info.Kind == SpProgramKind.External && IsRunning(row.Info.Id);
                if (running != row.IsRunning)
                {
                    row.IsRunning = running;
                    row.StateText = !row.IsAvailable
                        ? LanguageService.Get("SpPrograms_State_Todo")
                        : running
                            ? LanguageService.Get("SpPrograms_State_Running")
                            : LanguageService.Get("SpPrograms_State_Available");
                    changed = true;
                }
            }
            if (changed)
            {
                dg_programs.ItemsSource = null;
                dg_programs.ItemsSource = _rows;
            }
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
                txt_state.Text = !row.IsAvailable
                    ? string.Format(LanguageService.Get("SpPrograms_Msg_TodoFormat"), row.Name)
                    : row.Path;
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
            if (row.Info.Id == SpProgramCatalog.MoWorkOrderId)
            {
                WindowMesProductionQuery window = new(0) { Owner = this };
                window.ShowDialog();
                return;
            }
            if (row.Info.Id == SpProgramCatalog.PreIssueId)
            {
                WindowMesProductionQuery window = new(1) { Owner = this };
                window.ShowDialog();
                return;
            }
            if (row.Info.Id == SpProgramCatalog.InventoryId)
            {
                WindowMesProductionQuery window = new(2) { Owner = this };
                window.ShowDialog();
                return;
            }
            MessageBox.Show(this, string.Format(LanguageService.Get("SpPrograms_Msg_TodoFormat"), row.Name),
                LanguageService.Get("Cap_Info"), MessageBoxButton.OK, MessageBoxImage.Information);
        }

        /// <summary>
        /// 启动外部程式（它自己的窗口会弹出来）。
        /// 已经开着的就把那个窗口调到最前面；启动后**不关闭本窗口**，方便继续开别的或切回主窗口。
        /// </summary>
        private void OpenExternal(ProgramRow row)
        {
            if (IsRunning(row.Info.Id) && Running.TryGetValue(row.Info.Id, out Process running))
            {
                BringToFront(running);
                _logger.Info($"SP程序集：切换到已打开的 {row.Name}");
                return;
            }
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
                Process process = Process.Start(psi);
                if (process != null)
                {
                    Running[row.Info.Id] = process;
                }
                _logger.Info($"SP程序集：已启动 {row.Name}（{path}）");
                RefreshRunningState();
                // 把焦点交回刚启动的程式（它自己的窗口可能被本程序挡在后面）
                if (process != null)
                {
                    _ = process.WaitForInputIdle(3000);
                    BringToFront(process);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"SP程序集：启动 {path} 失败");
                MessageBox.Show(this, string.Format(LanguageService.Get("SpPrograms_Msg_LaunchFailedFormat"), ex.Message),
                    LanguageService.Get("Cap_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 把某个已打开的旧程式窗口调到最前（主窗口/本窗口随时可以把焦点抢回来，也随时能把它叫上来）。
        /// </summary>
        private static void BringToFront(Process process)
        {
            try
            {
                process.Refresh();
                IntPtr handle = process.MainWindowHandle;
                if (handle == IntPtr.Zero)
                {
                    return;
                }
                ShowWindow(handle, SwRestore);
                SetForegroundWindow(handle);
            }
            catch
            {
                // 旧程式窗口拿不到句柄时忽略（比如它还在启动中）
            }
        }

        private void Btn_Close_Click(object sender, RoutedEventArgs e) => Close();
    }
}

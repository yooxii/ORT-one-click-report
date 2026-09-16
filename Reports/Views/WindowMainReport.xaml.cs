using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Main.Views;
using ORT一键报告.Models;
using ORT一键报告.Reports.Models;
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
    /// <summary>
    /// WindowMainReport.xaml 的交互逻辑：
    /// 报告 Tab 按需显示（视图菜单勾选控制，懒加载，未勾选的 Tab 不占用内存）；
    /// 状态持久化到数据库，下次打开恢复。
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

        /* ###############################  按需 Tab 管理  ################################ */

        /// <summary>
        /// 当前已创建并加入 TabControl 的报告页面（按 ReportType 索引）
        /// </summary>
        private readonly Dictionary<string, (TabItem Tab, UserControl Page)> _tabs = [];

        /// <summary>
        /// 报告类型定义：显示名 → (菜单项获取函数, 页面创建函数, 默认勾选)
        /// </summary>
        private static readonly Dictionary<string, ReportTabDef> ReportTabDefs = new()
        {
            ["Thermal Shock"] = new("report.tab.thermalshock", () => new BaseReportPage { ReportType = "Thermal Shock", TestTime = 1 }, true),
            ["Burn In"] = new("report.tab.burnin", () => new BaseReportPage { ReportType = "Burn In", TestTime = 7 }, true),
            ["EMI"] = new("report.tab.emi", () => new EMIReportPage { ReportType = "EMI", TestTime = 1 }, false),
        };

        private sealed class ReportTabDef
        {
            public string SettingsKey { get; }
            public Func<UserControl> PageFactory { get; }
            public bool DefaultChecked { get; }

            public ReportTabDef(string settingsKey, Func<UserControl> pageFactory, bool defaultChecked)
            {
                SettingsKey = settingsKey;
                PageFactory = pageFactory;
                DefaultChecked = defaultChecked;
            }
        }

        /// <summary>
        /// 报告类型 → 对应菜单项的映射（懒初始化，InitMenuTabRefs 填充）
        /// </summary>
        private Dictionary<string, MenuItem> _menuByReport;

        public WindowMainReport()
        {
            InitializeComponent();

            ReportService = App.ServiceProvider.GetRequiredService<ReportService>();
            MainVM = App.ServiceProvider.GetRequiredService<MainReportViewModel>();
            DataContext = MainVM;

            _menuByReport = new Dictionary<string, MenuItem>
            {
                ["Thermal Shock"] = menu_tab_thermalshock,
                ["Burn In"] = menu_tab_burnin,
                ["EMI"] = menu_tab_emi,
            };

            ReportService.TemplateDir = Path.Combine(Directory.GetCurrentDirectory(), "Templates");
            ReportService.TempPath = Path.Combine(Path.GetTempPath(), "ORTTemp");

            // 先恢复 Tab，再预填数据（顺序重要：RestoreTabsFromSettings 必须在填充逻辑之前）
            Loaded += (s, e) =>
            {
                RestoreTabsFromSettings();
                FillFromPrefilledOnLoad();
            };
        }

        /// <summary>
        /// 从设置中恢复上次勾选的报告 Tab（懒加载，仅创建勾选的页面）
        /// </summary>
        private void RestoreTabsFromSettings()
        {
            AppSettingsService settings = App.ServiceProvider.GetRequiredService<AppSettingsService>();
            foreach (KeyValuePair<string, ReportTabDef> kv in ReportTabDefs)
            {
                bool isChecked = settings.GetBool(kv.Value.SettingsKey, kv.Value.DefaultChecked);
                if (_menuByReport.TryGetValue(kv.Key, out MenuItem menu))
                {
                    menu.IsChecked = isChecked;
                }
                if (isChecked)
                {
                    AddTab(kv.Key, kv.Value);
                }
            }
        }

        /// <summary>
        /// 创建并加入指定类型的报告 Tab（若已存在则直接返回）
        /// </summary>
        private void AddTab(string reportType, ReportTabDef def)
        {
            if (_tabs.ContainsKey(reportType))
            {
                return;
            }
            UserControl page = def.PageFactory();
            TabItem tab = new() { Header = reportType, Content = page };
            tab_report.Items.Add(tab);
            _tabs[reportType] = (tab, page);
            _logger.Info($"创建报告 Tab: {reportType}");
        }

        /// <summary>
        /// 移除指定类型的报告 Tab（若存在则释放页面）
        /// </summary>
        private void RemoveTab(string reportType)
        {
            if (_tabs.TryGetValue(reportType, out (TabItem Tab, UserControl Page) entry))
            {
                tab_report.Items.Remove(entry.Tab);
                _tabs.Remove(reportType);
                _logger.Info($"移除报告 Tab: {reportType}");
            }
        }

        /// <summary>
        /// 视图菜单勾选/取消时切换对应 Tab 的创建/释放，并持久化状态
        /// </summary>
        private void Menu_TabToggle_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not MenuItem menu)
            {
                return;
            }
            string reportType = _menuByReport.FirstOrDefault(kv => kv.Value == menu).Key;
            if (reportType == null || !ReportTabDefs.TryGetValue(reportType, out ReportTabDef def))
            {
                return;
            }
            AppSettingsService settings = App.ServiceProvider.GetRequiredService<AppSettingsService>();
            settings.SetBool(def.SettingsKey, menu.IsChecked);
            if (menu.IsChecked)
            {
                AddTab(reportType, def);
                if (IsLoaded && _tabs.TryGetValue(reportType, out (TabItem Tab, UserControl Page) entry))
                {
                    InitSinglePage(entry.Page);
                    // 勾选后才创建的 Tab 同样要按报告文件夹刷新表头、并填上携带的单体数据
                    if (ReportService.EnteredFromPlan)
                    {
                        ReadHeaderFromReportFolder(entry.Page);
                    }
                    if ((ReportService.UUTInfos?.SNs?.Count ?? 0) > 0)
                    {
                        FillDetailsForPage(entry.Page);
                    }
                }
            }
            else
            {
                RemoveTab(reportType);
            }
        }

        /* ###############################  加载/生成  ################################ */

        /// <summary>
        /// 窗口加载后：从预填数据填充表头与单体数据（必须在 RestoreTabsFromSettings 之后调用）
        /// </summary>
        private void FillFromPrefilledOnLoad()
        {
            try
            {
                foreach ((TabItem _, UserControl page) in _tabs.Values)
                {
                    InitSinglePage(page);
                }
                // 从计划表进入：表头与图片只以"该计划绑定的报告文件夹"里的本地报告为准，
                // 文件夹里没有这类报告就置空。不能用计划表的文案兜底 —— 那会把计划表里的
                // 机种/阶段/负责人/测试项目（如 "机型 + 测试项目"、"测试项目 (216hrs)"）当成
                // 报告表头呈现出来，而它们并不是这份报告里的信息。
                if (ReportService.EnteredFromPlan)
                {
                    foreach ((TabItem _, UserControl page) in _tabs.Values)
                    {
                        ReadHeaderFromReportFolder(page);
                    }
                }
                if (ReportService.UUTInfos != null && (ReportService.UUTInfos.SNs?.Count ?? 0) > 0)
                {
                    FillDetailsFromPrefilledUUT();
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "窗口加载预填数据失败");
            }
        }

        /// <summary>
        /// 初始化单个报告页面（列初始化 + 已有数据预填）
        /// </summary>
        private static void InitSinglePage(UserControl page)
        {
            if (page is BaseReportPage basePage)
            {
                basePage.InitReportPage();
            }
            // EMI 页面不需要 InitReportPage
        }

        /// <summary>
        /// 把预填的单体数据（序列号/工令/版本/周期）填进已打开 Tab 的明细表
        /// </summary>
        private void FillDetailsFromPrefilledUUT()
        {
            foreach ((TabItem _, UserControl page) in _tabs.Values)
            {
                FillDetailsForPage(page);
            }
            _logger.Info($"已从预填 UUTInfos 填充 {ReportService.UUTInfos.SNs?.Count ?? 0} 条单体数据");
        }

        /// <summary>
        /// 按"该计划绑定的报告文件夹"里的本地报告文件刷新单个报告页的表头与图片；
        /// 文件夹里没有这类报告时，页面内部会把表头与图片置空（不回退模板、不用计划表兜底）
        /// </summary>
        private static void ReadHeaderFromReportFolder(UserControl page)
        {
            switch (page)
            {
                case BaseReportPage basePage:
                    basePage.ReadReportHeader();
                    break;
                case EMIReportPage emiPage:
                    emiPage.ReadReportHeader();
                    break;
            }
        }

        /// <summary>
        /// 把预填的单体数据填进单个报告页的明细表（EMI 页没有明细表，跳过）
        /// </summary>
        private static void FillDetailsForPage(UserControl page)
        {
            if (page is BaseReportPage basePage)
            {
                basePage.SetReportResultData();
            }
        }

        private async void DoReport_Click(object sender, RoutedEventArgs e)
        {
            PopupWindow popup = new() { Title = LanguageService.Get("Title_Processing"), Message = "请耐心等待..." };
            if (sender is not Button btn)
            {
                return;
            }
            btn.IsEnabled = false;

            try
            {
                popup.Show();
                string ReportName = MainVM.ReportPath;
                if (string.IsNullOrWhiteSpace(ReportName) || !File.Exists(ReportName))
                {
                    throw new FileNotFoundException("报告概览文件不存在");
                }
                await MainVM.ReadInfoFromOverview(ReportName);
                _logger.Info("报告概览读取完成");

                // 仅对已打开的 Tab 执行读取与填充
                foreach ((TabItem _, UserControl page) in _tabs.Values)
                {
                    if (page is BaseReportPage basePage)
                    {
                        basePage.ReadReportHeader();
                        basePage.SetReportResultData();
                    }
                    else if (page is EMIReportPage emiPage)
                    {
                        emiPage.ReadReportHeader();
                    }
                }
                _logger.Info("表头数据已呈现至窗口");
            }
            catch (FileNotFoundException ex)
            {
                _logger.Error(ex, "报告文件不存在");
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_ReportFileNotFound"), LanguageService.Get("Cap_Error"));
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "读取报告出现错误");
                _ = MessageBox.Show($"读取报告出现错误{ex}", LanguageService.Get("Cap_Error"));
            }
            finally
            {
                popup.Close();
                btn.IsEnabled = true;
            }
        }

        private void MenuItem_ATE_Click(object sender, RoutedEventArgs e)
        {
            // 不设置 Owner：ATE 窗口最大化后会一直压住属主窗口，导致一键报告窗口无法切到前台
            ATEWindow ateWindow = new()
            {
                SubmitHandler = SubmitAteData
            };
            ateWindow.Show();
        }

        /// <summary>支持接收 ATE 数据（做成 ATE 报告并嵌入）的报告类型</summary>
        public static IReadOnlyList<string> AteReportTypes { get; } = ["Thermal Shock", "Burn In"];

        /// <summary>
        /// ATE 工具「提交到报告」：把生成好的 ATE 文件填到指定报告类型的页面里（Tab 未创建时先创建）。
        /// 报告生成时会把这个文件作为 OLE 附件嵌到报告里。
        /// </summary>
        public bool SubmitAteData(string reportType, string filePath)
        {
            if (string.IsNullOrWhiteSpace(reportType) || string.IsNullOrWhiteSpace(filePath)
                || !ReportTabDefs.TryGetValue(reportType, out ReportTabDef def))
            {
                return false;
            }
            AddTab(reportType, def);
            if (!_tabs.TryGetValue(reportType, out (TabItem Tab, UserControl Page) entry) || entry.Page is not BaseReportPage page)
            {
                return false;
            }
            page.SetAteData(filePath);
            tab_report.SelectedItem = entry.Tab;
            _logger.Info($"ATE 数据已提交到「{reportType}」报告：{filePath}");
            return true;
        }

        private void MenuItem_ReportTemplate_Click(object sender, RoutedEventArgs e)
        {
            WindowReportTemplate windowReportTemplate = new()
            {
                Owner = this
            };
            windowReportTemplate.Show();
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

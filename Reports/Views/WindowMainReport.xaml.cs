using Microsoft.Extensions.DependencyInjection;
using NLog;
using OfficeOpenXml;
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
            foreach ((TabItem _, UserControl page) in _tabs.Values)
            {
                InitSinglePage(page);
            }
            // 从计划表右键菜单携带记录打开时：预填表头 + 预填单体数据（无需等待读取报告概览）
            if (ReportService.MatchedPlan != null || ReportService.PrefilledReportModel != null)
            {
                ApplyMatchedPlanToAllTabs();
            }
            if (ReportService.UUTInfos != null && (ReportService.UUTInfos.SNs?.Count ?? 0) > 0)
            {
                FillDetailsFromPrefilledUUT();
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
        /// 从预填的 UUTInfos 填充已打开的 Thermal Shock / Burn In Tab 的单体数据（DataGrid）
        /// </summary>
        private void FillDetailsFromPrefilledUUT()
        {
            foreach ((TabItem _, UserControl page) in _tabs.Values)
            {
                if (page is BaseReportPage basePage)
                {
                    basePage.SetReportResultData();
                }
            }
            _logger.Info($"已从预填 UUTInfos 填充 {ReportService.UUTInfos.SNs?.Count ?? 0} 条单体数据");
        }

        private async void DoReport_Click(object sender, RoutedEventArgs e)
        {
            PopupWindow popup = new() { Title = "处理中", Message = "请耐心等待..." };
            if (sender is not Button btn)
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
                ApplyMatchedPlanToAllTabs();
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

        /// <summary>
        /// 对所有已打开的 BaseReportPage（Thermal Shock / Burn In）应用匹配的计划记录
        /// </summary>
        private void ApplyMatchedPlanToAllTabs()
        {
            foreach ((TabItem _, UserControl page) in _tabs.Values)
            {
                if (page is BaseReportPage basePage)
                {
                    ApplyMatchedPlanToHeader(basePage.ReportHeaderInfo);
                }
            }
        }

        /// <summary>
        /// 用领退和计划匹配到的记录补充报告表头（仅填充模板中为空的字段）。
        /// 若同时携带 PrefilledReportModel，其 Header 作为进一步兆底填充。
        /// </summary>
        private void ApplyMatchedPlanToHeader(ReportHeaderViewModel header)
        {
            Plan plan = ReportService.MatchedPlan;
            ORT一键报告.Reports.Models.ReportHeaderData prefilledHeader =
                ReportService.PrefilledReportModel?.Header;
            if ((plan == null && prefilledHeader == null) || header == null)
            {
                return;
            }
            // 优先级：MatchedPlan > PrefilledReportModel.Header
            if (string.IsNullOrWhiteSpace(header.PROJECT_NAME?.Data))
            {
                string value = plan?.TestItem != null ? $"{plan.ModelName} {plan.TestItem}" : (plan?.ModelName ?? prefilledHeader?.ProjectName);
                if (value != null)
                {
                    header.PROJECT_NAME ??= new DataCell();
                    header.PROJECT_NAME.Data = value;
                }
            }
            if (string.IsNullOrWhiteSpace(header.TEST_STAGE?.Data))
            {
                string value = plan?.Stage ?? prefilledHeader?.TestStage;
                if (value != null)
                {
                    header.TEST_STAGE ??= new DataCell();
                    header.TEST_STAGE.Data = value;
                }
            }
            if (string.IsNullOrWhiteSpace(header.TESTED_BY?.Data))
            {
                string value = plan?.Owner ?? prefilledHeader?.TestedBy;
                if (value != null)
                {
                    header.TESTED_BY ??= new DataCell();
                    header.TESTED_BY.Data = value;
                }
            }
            if (string.IsNullOrWhiteSpace(header.TestDescription?.Data))
            {
                string value = plan?.TestPeriod != null ? $"{plan.TestItem} ({plan.TestPeriod}hrs)"
                    : (plan?.TestItem ?? prefilledHeader?.TestDescription);
                if (value != null)
                {
                    header.TestDescription ??= new DataCell();
                    header.TestDescription.Data = value;
                }
            }
            // 携带的领退记录：S/N 等领退数据填入仍为空的相应位置（尽力填充，不覆盖已有内容）
            Requisition req = ReportService.MatchedRequisition;
            if (req != null)
            {
                if (string.IsNullOrWhiteSpace(header.TestDescription?.Data) && !string.IsNullOrWhiteSpace(req.SN))
                {
                    header.TestDescription ??= new DataCell();
                    header.TestDescription.Data = $"S/N: {req.SN}";
                }
            }
            if (plan != null)
            {
                _logger.Info($"已用匹配计划记录(Id={plan.Id})补充报告表头信息");
            }
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

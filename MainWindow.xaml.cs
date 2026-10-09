using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Admin.Views;
using ORT一键报告.Main.Views;
using ORT一键报告.Models;
using ORT一键报告.Plans.Views;
using ORT一键报告.Reports.Views;
using ORT一键报告.Review.Views;
using ORT一键报告.Services;
using ORT一键报告.Utils;
using ORT一键报告.ViewModels;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using DrawingIcon = System.Drawing.Icon;
using DrawingSystemIcons = System.Drawing.SystemIcons;
using WinForms = System.Windows.Forms;

namespace ORT一键报告
{

    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        private readonly AuthService _auth;
        private readonly IPermissionService _permission;
        private readonly ReviewService _reviewService;
        private readonly AppSettingsService _appSettings;
        private readonly MailNotifier _mailNotifier;
        private readonly PlanIndexScheduler _indexScheduler;

        /// <summary>计划结束日期提醒的定时检查（每 6 小时一次）</summary>
        private readonly System.Windows.Threading.DispatcherTimer _mailTimer = new()
        {
            Interval = TimeSpan.FromHours(6)
        };

        /// <summary>右下角托盘图标（最小化到后台时显示；退出或关闭后台运行后释放）</summary>
        private WinForms.NotifyIcon _trayIcon;

        /// <summary>托盘菜单当前使用的字体（每次重建都会换一个，旧的在重建时释放）</summary>
        private System.Drawing.Font _trayMenuFont;

        /// <summary>本次关闭是否真的要退出程序（托盘「退出程序」/ 注销登录 / 系统结束会话时为 true）</summary>
        private bool _exitRequested;

        /// <summary>当前是否已最小化到后台（窗口隐藏、托盘图标有效）</summary>
        private bool _inBackground;

        /// <summary>本次运行是否以「后台运行」状态启动（由 --background 命令行参数决定）</summary>
        public bool StartInBackground { get; set; }

        public MainViewModel MainVM { get; set; }

        public MainWindow()
        {
            InitializeComponent();

            _auth = App.ServiceProvider.GetRequiredService<AuthService>();
            _permission = App.ServiceProvider.GetRequiredService<IPermissionService>();
            _reviewService = App.ServiceProvider.GetRequiredService<ReviewService>();
            _appSettings = App.ServiceProvider.GetRequiredService<AppSettingsService>();
            _mailNotifier = App.ServiceProvider.GetRequiredService<MailNotifier>();
            _indexScheduler = App.ServiceProvider.GetRequiredService<PlanIndexScheduler>();

            MainVM = App.ServiceProvider.GetRequiredService<MainViewModel>();
            DataContext = MainVM;
            MainVM.SubscribeLanguageChange();
            // 语言切换时刷新菜单文案与左下角用户身份信息
            LanguageService.LanguageChanged += () => Dispatcher.Invoke(UpdateUIByPermission);

            Loaded += (s, e) => Activate();
            // 启动时应用设置字体，并在设置变更时实时刷新所有已打开窗口
            Loaded += (s, e) => _appSettings.ApplyFont(this);
            _appSettings.SettingsChanged += () => Dispatcher.Invoke(() => _appSettings.ApplyFontToAll());
            // 主界面背景图（设置→界面 里选择）
            Loaded += (s, e) => ApplyBackgroundImage();
            _appSettings.SettingsChanged += () => Dispatcher.Invoke(ApplyBackgroundImage);

            _auth.AuthChanged += () => Dispatcher.Invoke(UpdateUIByPermission);
            Loaded += (s, e) => UpdateUIByPermission();
            Loaded += (s, e) => StartMailReminder();
            Loaded += (s, e) => StartPlanIndexScheduler();
            Closed += (s, e) => _mailTimer.Stop();
            // 关闭主窗口：先问是「最小化到后台」还是「退出程序」（见 OnClosing）
            Closing += MainWindow_Closing;
            Closed += (s, e) => DisposeTrayIcon();
        }

        /// <summary>
        /// 系统结束会话 / 注销登录：不再拦截关闭，直接放行（由 App 的 SessionEnding 调用）
        /// </summary>
        public void AllowExitForSessionEnding() => _exitRequested = true;

        /* ###############################  后台运行（托盘）  ################################ */

        /// <summary>
        /// 把窗口收进右下角托盘：隐藏窗口、显示托盘图标，并顺手清理一次内存；
        /// 程序本身继续运行（报告扫描、计划索引、邮件提醒等后台任务照常）。
        /// </summary>
        public void MinimizeToBackground()
        {
            try
            {
                Hide();
                ShowInTaskbar = false;
                EnsureTrayIcon();
                _inBackground = true;
                // 界面相关的托管对象（图片、表格、报告预览等）暂时用不到了，回收并把工作集还给系统
                MemoryTrimmer.Trim("最小化到后台");
                _logger.Info("已最小化到后台（托盘）");
                ShowBackgroundTipIfDue();
            }
            catch (Exception ex)
            {
                _logger.Warn($"最小化到后台失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 最小化到后台后弹一次托盘气球提示，告知用户可在设置里关掉这个行为；
        /// 「每天只提醒一次」：记住上次弹出的日期，同一天再最小化就不打扰了。
        ///
        /// 刻意**延迟**约 1.5 秒再弹：第一次最小化时托盘图标是刚创建的，
        /// 立刻弹气球会抢走刚落到图标上的鼠标交互——用户这时右键点不开菜单、双击也像没反应。
        /// 等图标在托盘里稳定下来再提示，鼠标就能正常操作。
        /// </summary>
        private void ShowBackgroundTipIfDue()
        {
            if (_trayIcon == null)
            {
                return;
            }
            string today = DateTime.Today.ToString("yyyy-MM-dd");
            if (string.Equals(_appSettings.GetLastTrayTipDate(), today, StringComparison.Ordinal))
            {
                return;
            }
            // 先把日期记下，避免提示排队期间反复最小化导致弹多次
            _appSettings.SetLastTrayTipDate(today);

            System.Windows.Threading.DispatcherTimer timer = new()
            {
                Interval = TimeSpan.FromMilliseconds(1500)
            };
            timer.Tick += (s, e) =>
            {
                timer.Stop();
                if (_trayIcon == null)
                {
                    return;
                }
                try
                {
                    _trayIcon.ShowBalloonTip(5000,
                        LanguageService.Get("Main_TrayBalloonTitle"),
                        LanguageService.Get("Main_TrayBalloonText"),
                        WinForms.ToolTipIcon.Info);
                    _logger.Info($"已提示最小化到后台（{today}，当天不再重复提示）");
                }
                catch (Exception ex)
                {
                    _logger.Warn($"弹出最小化提示失败: {ex.Message}");
                }
            };
            timer.Start();
        }

        /// <summary>
        /// 从托盘恢复主窗口（托盘菜单「显示主窗口」、重复启动程序时走这里）
        /// </summary>
        public void RestoreFromBackground()
        {
            try
            {
                _inBackground = false;
                if (!IsVisible)
                {
                    Show();
                }
                ShowInTaskbar = true;
                if (WindowState == WindowState.Minimized)
                {
                    WindowState = WindowState.Normal;
                }
                Activate();
                Topmost = true;
                Topmost = false;
                Focus();
                _logger.Info("已从后台恢复主窗口");
            }
            catch (Exception ex)
            {
                _logger.Warn($"恢复主窗口失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 真正退出程序：置退出标记并关窗（OnClosing 不再拦截，OnClosed 里统一 Shutdown）
        /// </summary>
        public void ExitApplication()
        {
            _exitRequested = true;
            Close();
        }

        /// <summary>
        /// 主窗口关闭：按设置里的「最小化到后台」直接决定行为——
        /// 勾选＝收进托盘继续运行（取消本次关闭），未勾选＝直接退出程序，都不再弹询问框。
        /// 托盘「退出程序」、注销、系统结束会话这三种情况始终直接退出。
        /// </summary>
        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            if (_exitRequested || Application.Current == null)
            {
                return;
            }
            if (App.IsExiting)
            {
                _exitRequested = true;
                return;
            }
            if (!_appSettings.MinimizeToTrayOnClose)
            {
                _exitRequested = true;
                return;
            }
            e.Cancel = true;
            MinimizeToBackground();
        }

        /// <summary>
        /// 按需创建托盘图标（最小化到后台时才需要；用户关掉后台运行后不再创建）。
        /// 只保留右键菜单这一个入口：双击图标不再打开主窗口（避免误触把窗口弹出来）。
        /// </summary>
        private void EnsureTrayIcon()
        {
            if (_trayIcon != null)
            {
                return;
            }
            _trayIcon = new WinForms.NotifyIcon
            {
                Icon = LoadTrayIcon(),
                Visible = true
            };
            BuildTrayMenu();
            // 设置里改字体/语言/开机自启后托盘菜单跟着刷新
            _appSettings.SettingsChanged += OnSettingsChangedForTray;
        }

        /* ###############################  托盘菜单  ################################ */

        /// <summary>托盘菜单分隔线的标记（ToolStripItem 只能属于一个菜单，因此每次都新建，靠这个标记识别）</summary>
        private const string TraySeparatorTag = "tray-separator";

        /// <summary>
        /// 重建整个托盘菜单：按当前权限禁用无权使用的项，并同步开机自启的勾选状态。
        /// 语言、权限、开机自启状态变化后都走这里重建，避免按下标改文字改错项。
        /// </summary>
        private void BuildTrayMenu()
        {
            if (_trayIcon == null)
            {
                return;
            }
            try
            {
                DisposeTrayMenu();
                WinForms.ContextMenuStrip menu = new();
                _trayMenuFont = TrayMenuFont();
                menu.Font = _trayMenuFont;

                AddTrayItem(menu.Items, LanguageService.Get("Main_TrayShow"), true, (s, e) => RestoreFromBackground());
                AddTraySeparator(menu.Items);

                // 领用与计划：打开需要浏览权限；新增/批量回线/回线转移单还需要编辑权限
                WinForms.ToolStripMenuItem plans = new(LanguageService.Get("MainWindow_Button_Plans"))
                {
                    Enabled = _permission.Can("plan.view")
                };
                menu.Items.Add(plans);
                AddTrayItem(plans.DropDownItems, LanguageService.Get("Plans_Menu_OpenRequisitions"), true, (s, e) => OpenPlans());
                AddTrayItem(plans.DropDownItems, LanguageService.Get("Plans_Menu_OpenPlans"), true, (s, e) => OpenPlans());
                AddTraySeparator(plans.DropDownItems);
                AddTrayItem(plans.DropDownItems, LanguageService.Get("Plans_AddRequisition"), CanPlanEdit, (s, e) => OpenPlans(() => InvokeOpenPlans(w => w.InvokeAddRequisition())));
                AddTrayItem(plans.DropDownItems, LanguageService.Get("Plans_AddPlan"), CanPlanEdit, (s, e) => OpenPlans(() => InvokeOpenPlans(w => w.InvokeAddPlan())));
                AddTraySeparator(plans.DropDownItems);
                AddTrayItem(plans.DropDownItems, LanguageService.Get("Plans_Menu_BatchReturn"), CanPlanEdit, (s, e) => OpenPlans(() => InvokeOpenPlans(w => w.InvokeBatchReturn())));
                AddTrayItem(plans.DropDownItems, LanguageService.Get("Plans_ReturnTransfer"), CanPlanEdit, (s, e) => OpenPlans(() => InvokeOpenPlans(w => w.InvokeReturnTransfer())));

                // 一键报告：Burn In / Thermal Shock / EMI 三个报告页 + ATE 工具
                bool canReport = _permission.Can("report.use");
                WinForms.ToolStripMenuItem report = new(LanguageService.Get("MainWindow_Button_Report"))
                {
                    Enabled = canReport
                };
                menu.Items.Add(report);
                AddTrayItem(report.DropDownItems, LanguageService.Get("Main_TrayReportBurnIn"), true, (s, e) => OpenReport("Burn In"));
                AddTrayItem(report.DropDownItems, LanguageService.Get("Main_TrayReportThermalShock"), true, (s, e) => OpenReport("Thermal Shock"));
                AddTrayItem(report.DropDownItems, LanguageService.Get("Main_TrayReportEmi"), true, (s, e) => OpenReport("EMI"));
                AddTraySeparator(report.DropDownItems);
                AddTrayItem(report.DropDownItems, LanguageService.Get("Main_TrayAteTool"), true, (s, e) => OpenAteTool());

                // 管理：打开管理表
                WinForms.ToolStripMenuItem admin = new(LanguageService.Get("MainWindow_Button_Admin"))
                {
                    Enabled = _permission.Can("admin.manage")
                };
                menu.Items.Add(admin);
                AddTrayItem(admin.DropDownItems, LanguageService.Get("Main_TrayAdminOpen"), true, (s, e) => OpenAdmin());

                // 审查
                AddTrayItem(menu.Items, LanguageService.Get("MainWindow_Button_Review"), _permission.Can("review.view"), (s, e) => OpenReview());

                AddTraySeparator(menu.Items);

                // 设置：打开设置 + 开机自启开关（远程路径下不可自启，直接禁用）
                WinForms.ToolStripMenuItem settings = new(LanguageService.Get("MainWindow_Menu_Settings"));
                menu.Items.Add(settings);
                AddTrayItem(settings.DropDownItems, LanguageService.Get("Main_TrayOpenSettings"), true, (s, e) => OpenSettings());
                bool autoStartAvailable = ShouldOfferAutoStart();
                AddTrayItem(settings.DropDownItems, LanguageService.Get("Main_TrayAutoStart"), autoStartAvailable, (s, e) => ToggleAutoStart())
                    .Checked = autoStartAvailable && StartupManager.IsEnabled();

                AddTraySeparator(menu.Items);
                AddTrayItem(menu.Items, LanguageService.Get("Main_TrayExit"), true, (s, e) => ExitApplication());

                _trayIcon.ContextMenuStrip = menu;
                UpdateTrayTip();
            }
            catch (Exception ex)
            {
                _logger.Warn($"构建托盘菜单失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 往菜单里加一项，返回新加的项以便继续设置属性
        /// </summary>
        private static WinForms.ToolStripMenuItem AddTrayItem(WinForms.ToolStripItemCollection items, string text, bool enabled, EventHandler onClick)
        {
            WinForms.ToolStripMenuItem item = new(text) { Enabled = enabled };
            if (onClick != null)
            {
                item.Click += onClick;
            }
            items.Add(item);
            return item;
        }

        /// <summary>往菜单里加一条分隔线</summary>
        private static void AddTraySeparator(WinForms.ToolStripItemCollection items)
            => items.Add(new WinForms.ToolStripSeparator { Tag = TraySeparatorTag });

        /// <summary>托盘菜单字体跟随设置里的界面字体与字号（夹在 8–12 之间，托盘菜单过大会很怪）</summary>
        private System.Drawing.Font TrayMenuFont()
        {
            try
            {
                return new System.Drawing.Font(
                    _appSettings.Settings.UI.FontFamily,
                    (float)Math.Max(8, Math.Min(12, _appSettings.Settings.UI.FontSize)));
            }
            catch (Exception ex)
            {
                _logger.Warn($"托盘菜单字体回退为默认: {ex.Message}");
                return System.Drawing.SystemFonts.MenuFont;
            }
        }

        /// <summary>释放上一次构建的菜单（连同菜单里自己建的项与字体）</summary>
        private void DisposeTrayMenu()
        {
            // 每次重建都会新建字体，这里一并释放，避免反复开关窗口攒下一堆字体句柄
            try
            {
                _trayMenuFont?.Dispose();
            }
            catch (Exception ex)
            {
                _logger.Warn($"释放托盘菜单字体失败: {ex.Message}");
            }
            _trayMenuFont = null;

            WinForms.ContextMenuStrip menu = _trayIcon?.ContextMenuStrip;
            if (menu == null)
            {
                return;
            }
            _trayIcon.ContextMenuStrip = null;
            foreach (WinForms.ToolStripItem item in menu.Items)
            {
                if (!Equals(item.Tag, TraySeparatorTag))
                {
                    item.Dispose();
                }
            }
            menu.Dispose();
        }

        /// <summary>
        /// 在已打开的领退和计划窗口上执行动作；窗口有可能刚被用户关掉，这里兜底重开一次
        /// </summary>
        private static void InvokeOpenPlans(Action<Plans.Views.WindowPlans> action)
        {
            foreach (Window w in Application.Current.Windows)
            {
                if (w is Plans.Views.WindowPlans plans)
                {
                    action(plans);
                    return;
                }
            }
            // 窗口在点菜单的瞬间被关掉了：重开一个再执行
            Plans.Views.WindowPlans reopened = new();
            reopened.Show();
            action(reopened);
        }

        /// <summary>
        /// 是否提供「开机自启」开关：程序位于远程路径（网络共享/网络盘/SUBST 虚拟盘）时不可自启，
        /// 按需求直接禁用该菜单项
        /// </summary>
        private static bool ShouldOfferAutoStart() => !StartupManager.IsRemoteLocation();

        /// <summary>托盘菜单「开机自启」：在开与关之间切换（远程路径下菜单项已禁用，这里再挡一次）</summary>
        private void ToggleAutoStart()
        {
            if (StartupManager.TryDescribeRemoteLocation() is string remote)
            {
                _ = MessageBox.Show(
                    string.Format(LanguageService.Get("Settings_AutoStartRemoteBlockedFormat"), remote),
                    LanguageService.Get("Cap_Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            bool enable = !StartupManager.IsEnabled();
            BackgroundSettings background = _appSettings.GetBackgroundSettings();
            bool toBackground = enable && background.AutoStartToBackground;
            if (!StartupManager.TryEnable(enable, toBackground, out string error))
            {
                _ = MessageBox.Show(
                    string.Format(LanguageService.Get("Settings_AutoStartFailedFormat"), error),
                    LanguageService.Get("Cap_Error"), MessageBoxButton.OK, MessageBoxImage.Warning);
                BuildTrayMenu();
                return;
            }
            // 同步回本机设置，设置窗口里的勾选状态跟着变
            background.AutoStart = enable;
            background.AutoStartToBackground = toBackground;
            _appSettings.SetBackgroundSettings(background);
            _logger.Info(enable ? "托盘菜单：已开启开机自启" : "托盘菜单：已关闭开机自启");
            BuildTrayMenu();
        }

        /// <summary>
        /// 刷新托盘图标的提示文字（悬停显示）
        /// </summary>
        private void UpdateTrayTip()
        {
            if (_trayIcon == null)
            {
                return;
            }
            try
            {
                _trayIcon.Text = LanguageService.Get("Main_TrayTip");
            }
            catch (Exception ex)
            {
                _logger.Warn($"刷新托盘提示失败: {ex.Message}");
            }
        }

        /// <summary>刷新托盘提示与菜单（语言、字体、权限、开机自启状态变化后调用）</summary>
        private void UpdateTrayText()
        {
            UpdateTrayTip();
            BuildTrayMenu();
        }

        private void OnSettingsChangedForTray()
        {
            Dispatcher.Invoke(() =>
            {
                if (_trayIcon == null)
                {
                    return;
                }
                if (!_appSettings.MinimizeToTrayOnClose && !_inBackground)
                {
                    // 用户取消了「关闭时最小化到后台」：清掉托盘图标，不留下无用的后台入口
                    DisposeTrayIcon();
                    return;
                }
                UpdateTrayText();
            });
        }

        /// <summary>
        /// 释放托盘图标（程序退出、或用户关掉后台运行时调用）
        /// </summary>
        public void DisposeTrayIcon()
        {
            if (_trayIcon == null)
            {
                return;
            }
            try
            {
                _trayIcon.Visible = false;
                DisposeTrayMenu();
                _trayIcon.Dispose();
            }
            catch (Exception ex)
            {
                _logger.Warn($"释放托盘图标失败: {ex.Message}");
            }
            finally
            {
                _trayIcon = null;
                _appSettings.SettingsChanged -= OnSettingsChangedForTray;
            }
        }

        /// <summary>
        /// 托盘图标：优先取程序自带图标；取不到时退回系统信息图标（不影响功能）
        /// </summary>
        private static DrawingIcon LoadTrayIcon()
        {
            try
            {
                string exe = StartupManager.ProgramPath();
                if (!string.IsNullOrEmpty(exe) && System.IO.File.Exists(exe))
                {
                    DrawingIcon icon = System.Drawing.Icon.ExtractAssociatedIcon(exe);
                    if (icon != null)
                    {
                        return icon;
                    }
                }
            }
            catch
            {
                // 忽略，用系统图标兜底
            }
            return DrawingSystemIcons.Application;
        }

        /// <summary>
        /// 应用主界面背景图：从设置里取路径，铺满窗口并显示淡淡遮罩；
        /// 路径为空或文件不存在时回到主题背景。
        /// </summary>
        private void ApplyBackgroundImage()
        {
            string path = _appSettings.Settings?.UI?.BackgroundImage;
            if (!string.IsNullOrWhiteSpace(path) && System.IO.File.Exists(path))
            {
                try
                {
                    System.Windows.Media.Imaging.BitmapImage image = new();
                    image.BeginInit();
                    image.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad; // 读完即释放，不锁文件
                    image.UriSource = new Uri(path, UriKind.Absolute);
                    image.EndInit();
                    img_background.Source = image;
                    img_background.Visibility = Visibility.Visible;
                    rect_scrim.Visibility = Visibility.Visible;
                    _logger.Info($"主界面背景图已应用：{path}");
                    return;
                }
                catch (Exception ex)
                {
                    _logger.Warn($"加载主界面背景图失败：{ex.Message}");
                }
            }
            img_background.Source = null;
            img_background.Visibility = Visibility.Collapsed;
            rect_scrim.Visibility = Visibility.Collapsed;
        }

        /// <summary>
        /// 启动"空闲时自动建立计划索引"：设置里开启（且为管理员）时才会真正执行；
        /// 任务与进度都落库，关掉程序或换台电脑都会从断点继续。
        /// </summary>
        private void StartPlanIndexScheduler()
        {
            try
            {
                _indexScheduler.Finished += message =>
                {
                    if (!string.IsNullOrWhiteSpace(message))
                    {
                        ToastService.Show(message, ORT一键报告.Main.Views.ToastType.Info);
                    }
                };
                _indexScheduler.Start();
            }
            catch (Exception ex)
            {
                _logger.Warn($"启动空闲计划索引失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 启动后台邮件提醒：延迟首次检查（避免拖慢启动），之后每 6 小时检查一次
        /// </summary>
        private async void StartMailReminder()
        {
            try
            {
                _mailTimer.Tick += (s, e) => _mailNotifier.CheckPlanDeadlinesInBackground();
                _mailTimer.Start();
                await System.Threading.Tasks.Task.Delay(TimeSpan.FromSeconds(10));
                _mailNotifier.CheckPlanDeadlinesInBackground();
            }
            catch (Exception ex)
            {
                _logger.Warn($"启动邮件提醒失败: {ex.Message}");
            }
        }

        /* ###############################  功能函数  ################################ */

        /// <summary>
        /// 主窗口关闭时一并关闭所有子窗口（退出程序）
        /// </summary>
        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            Application.Current.Shutdown();
        }

        /// <summary>
        /// 是否有领退表编辑权限（托盘菜单「新增领用/新增计划/批量回线/回线转移单」的可用性判据）
        /// </summary>
        private bool CanPlanEdit => _permission.Can("plan.edit");

        /// <summary>
        /// 根据当前登录状态与角色刷新入口可用性，并关闭当前无权限访问的子窗口
        /// </summary>
        private void UpdateUIByPermission()
        {
            // 菜单项只显示登录/注销动作，当前用户身份信息统一在窗口左下角展示
            menu_account.Header = _auth.CurrentUser == null
                ? LanguageService.Get("Main_Login")
                : LanguageService.Get("Main_Logout");
            // 用户中心仅登录后可用（游客无个人资料可管理）
            menu_user_center.IsEnabled = _auth.CurrentUser != null;
            txt_user_identity.Text = string.Format(LanguageService.Get("Main_IdentityFormat"), _auth.CurrentDisplayName);
            // 程序名下方的版本号：跟随程序集版本（AssemblyInfo 的 AssemblyVersion，随生成自动更新、不写死）。
            // build 段非 0 显示三段（如 0.3.2），否则两段（如 0.4），与更新日志「## 主.次」写法保持一致
            System.Version ver = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            string verText = ver == null ? "-" : ver.Build > 0 ? ver.ToString(3) : ver.ToString(2);
            txt_version.Text = string.Format(LanguageService.Get("Main_VersionFormat"), verText);
            btn_report.IsEnabled = _permission.Can("report.use");
            btn_admin.IsEnabled = _permission.Can("admin.manage");
            btn_review.IsEnabled = _permission.Can("review.view");
            btn_review.Content = _permission.Can("review.view")
                ? string.Format(LanguageService.Get("Main_ReviewCountFormat"), _reviewService.PendingCount())
                : LanguageService.Get("Main_Review");

            // 权限变化时关闭当前无权限访问的子窗口
            CloseUnauthorizedWindows();
            // 托盘菜单里的项也按权限决定可用性：登录/注销/换人后重建，避免留下无权使用的入口
            BuildTrayMenu();
        }

        /// <summary>
        /// 「用户」按钮：点击弹出上下文菜单（用户中心 / 登录-注销）
        /// </summary>
        private void Button_User_Click(object sender, RoutedEventArgs e)
        {
            if (btn_user.ContextMenu != null)
            {
                btn_user.ContextMenu.PlacementTarget = btn_user;
                btn_user.ContextMenu.IsOpen = true;
            }
        }

        /// <summary>
        /// 关闭当前登录状态无权限访问的子窗口。
        /// 游客：关闭领退和计划 / 一键报告 / 管理 / 审核；
        /// 已登录：按权限关闭对应子窗口。
        /// </summary>
        private void CloseUnauthorizedWindows()
        {
            List<Window> toClose = [];
            foreach (Window w in Application.Current.Windows)
            {
                if (w == this)
                {
                    continue;
                }
                switch (w)
                {
                    case Plans.Views.WindowPlans when !_permission.Can("plan.view") || _auth.CurrentUser == null:
                        toClose.Add(w);
                        break;
                    case Reports.Views.WindowMainReport when !_permission.Can("report.use"):
                        toClose.Add(w);
                        break;
                    case Admin.Views.WindowAdmin when !_permission.Can("admin.manage"):
                        toClose.Add(w);
                        break;
                    case Review.Views.WindowReview when !_permission.Can("review.view"):
                        toClose.Add(w);
                        break;
                    case Main.Views.WindowUserCenter when _auth.CurrentUser == null:
                        toClose.Add(w); // 注销后用户中心无内容可管理
                        break;
                }
            }
            foreach (Window w in toClose)
            {
                try
                {
                    w.Close();
                }
                catch (Exception ex)
                {
                    _logger.Warn($"关闭无权限窗口失败: {w.GetType().Name}, {ex.Message}");
                }
            }
        }

        /* ###############################  事件函数  ################################ */

        private void MenuItem_Login_Click(object sender, RoutedEventArgs e)
        {
            if (_auth.CurrentUser != null)
            {
                // 已登录 → 注销
                if (MessageBox.Show(string.Format(LanguageService.Get("Msg_ConfirmLogout"), _auth.CurrentDisplayName), LanguageService.Get("Cap_LogoutConfirm"),
                    MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    _auth.Logout();
                }
                return;
            }
            WindowLogin loginWindow = new()
            {
            };
            if (loginWindow.ShowDialog() == true)
            {
                _logger.Info($"当前用户: {_auth.CurrentDisplayName}");
                PromptCompleteEmailIfNeeded();
                PromptSetPasswordIfNeeded();
            }
        }

        /// <summary>
        /// 用户中心：管理当前登录用户自己的资料（显示名/邮箱/密码、本机登录信息）
        /// </summary>
        private void MenuItem_UserCenter_Click(object sender, RoutedEventArgs e)
        {
            if (_auth.CurrentUser == null)
            {
                _ = MessageBox.Show(LanguageService.Get("UserCenter_Msg_NeedLogin"), LanguageService.Get("Cap_NoPermission"));
                return;
            }
            // 已打开的则聚焦，不重复打开
            foreach (Window w in Application.Current.Windows)
            {
                if (w is WindowUserCenter existing)
                {
                    existing.Activate();
                    return;
                }
            }
            WindowUserCenter window = new();
            window.Show();
        }

        /// <summary>
        /// 账号还没设置密码（登录时只给了用户名）→ 提示去用户中心设置密码
        /// </summary>
        private void PromptSetPasswordIfNeeded()
        {
            if (!_auth.NeedsPasswordSetup || _auth.CurrentUser == null)
            {
                return;
            }
            if (MessageBox.Show(string.Format(LanguageService.Get("Msg_SetPasswordPromptFormat"), _auth.CurrentUser.Username),
                LanguageService.Get("Dlg_SetPasswordTitle"), MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            {
                return;
            }
            MenuItem_UserCenter_Click(this, new RoutedEventArgs());
        }

        /// <summary>
        /// 技术员/审核员登录后若未填写邮箱，提示完善（可跳过）
        /// </summary>
        private void PromptCompleteEmailIfNeeded()
        {
            if (!_auth.NeedsEmailCompletion)
            {
                return;
            }
            string title = LanguageService.Get("Dlg_EmailTitle");
            string prompt = string.Format(LanguageService.Get("Msg_EmailPromptFormat"), _auth.CurrentDisplayName);
            if (MessageBox.Show(prompt, title, MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            {
                return;
            }
            while (true)
            {
                WindowAdminInput input = new(title,
                    (LanguageService.Get("Admin_Email"), _auth.CurrentUser?.Email ?? "", false))
                {
                };
                if (input.ShowDialog() != true)
                {
                    return;
                }
                string email = input.Values[0];
                if (!AuthService.IsValidEmail(email))
                {
                    _ = MessageBox.Show(LanguageService.Get("Msg_EmailInvalid"), LanguageService.Get("Cap_Info"));
                    continue;
                }
                if (_auth.SetCurrentUserEmail(email))
                {
                    _ = MessageBox.Show(LanguageService.Get("Msg_EmailSaved"), LanguageService.Get("Cap_Success"));
                }
                return;
            }
        }

        private void MenuItem_ViewLog_Click(object sender, RoutedEventArgs e)
        {
            WindowLog windowLog = new()
            {
            };
            windowLog.Show();
        }

        /// <summary>
        /// 帮助→更新日志：只读展示随程序发布的「更新日志.md」（已打开则聚焦，不重复打开）
        /// </summary>
        private void MenuItem_Changelog_Click(object sender, RoutedEventArgs e)
        {
            foreach (Window w in Application.Current.Windows)
            {
                if (w is WindowChangelog existing)
                {
                    existing.Activate();
                    return;
                }
            }
            WindowChangelog window = new();
            window.Show();
        }

        /// <summary>
        /// 工具菜单：报告模板工具（由机种测试计划直接生成报告模板）
        /// </summary>
        private void MenuItem_ReportTemplate_Click(object sender, RoutedEventArgs e)
        {
            WindowReportTemplate window = new();
            window.Show();
        }

        private void MenuItem_Settings_Click(object sender, RoutedEventArgs e) => OpenSettings();

        /// <summary>打开设置窗口（主窗口菜单与托盘菜单共用）</summary>
        internal void OpenSettings()
        {
            WindowAppSettings settingsWindow = new()
            {
                Owner = null
            };
            settingsWindow.Show();
        }

        private void MenuItem_Quit_Click(object sender, RoutedEventArgs e)
        {
            ExitApplication();
        }

        private void Button_YiJianBaoGao_Click(object sender, RoutedEventArgs e) => OpenReport();

        /// <summary>
        /// 打开一键报告窗口；<paramref name="reportType"/> 非空时（托盘菜单：Burn In / Thermal Shock / EMI）
        /// 还要切到对应报告页，且预填与刷新要等窗口加载完再做
        /// </summary>
        internal void OpenReport(string reportType = null)
        {
            if (!_permission.Can("report.use"))
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_ReportNeedLogin"), LanguageService.Get("Cap_NoPermission"));
                return;
            }
            ToastService.WarnIfReportPathEmpty();
            // 直接进入：清掉上一次从计划表进入残留的匹配记录（ReportService 是单例，
            // 否则概览读到的 SN/工令/版本会被上一次的领用表数据覆盖）
            App.ServiceProvider.GetRequiredService<ReportService>().ClearMatchedSource();
            // 指定报告页时先显示窗口：页面创建要读表头、填单体数据，得等窗口加载完成
            WindowMainReport windowMainReport = new();
            if (string.IsNullOrEmpty(reportType))
            {
                windowMainReport.Show();
                return;
            }
            windowMainReport.Loaded += (s, args) => windowMainReport.SelectReportTab(reportType);
            windowMainReport.Show();
        }

        /// <summary>
        /// 打开 ATE 工具（托盘菜单「一键报告 → ATE 工具」）
        /// </summary>
        internal void OpenAteTool()
        {
            if (!_permission.Can("report.use"))
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_ReportNeedLogin"), LanguageService.Get("Cap_NoPermission"));
                return;
            }
            new ATEWindow().Show();
        }

        private void Button_Plans_Click(object sender, RoutedEventArgs e) => OpenPlans();

        /// <summary>
        /// 打开领退和计划窗口（已打开则聚焦）。
        /// <paramref name="pendingAction"/> 非空时（托盘菜单：新增领用/新增计划/批量回线/回线转移单），
        /// 等窗口首次显示完成后自动执行该动作。
        /// </summary>
        internal void OpenPlans(Action pendingAction = null)
        {
            Plans.Views.WindowPlans existing = null;
            foreach (Window w in Application.Current.Windows)
            {
                if (w is Plans.Views.WindowPlans found)
                {
                    existing = found;
                    break;
                }
            }
            if (existing != null)
            {
                if (existing.WindowState == WindowState.Minimized)
                {
                    existing.WindowState = WindowState.Normal;
                }
                existing.Activate();
                // 窗口之前就开着：直接执行（新增/批量窗口都是即时动作，不依赖刚加载完）
                pendingAction?.Invoke();
                return;
            }
            Plans.Views.WindowPlans windowPlans = new(pendingAction);
            windowPlans.Show();
        }

        private void Button_Admin_Click(object sender, RoutedEventArgs e) => OpenAdmin();

        /// <summary>打开管理窗口（托盘菜单「管理 → 打开管理表」）</summary>
        internal void OpenAdmin()
        {
            if (!_permission.Can("admin.manage"))
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_AdminNeedLogin"), LanguageService.Get("Cap_NoPermission"));
                return;
            }
            WindowAdmin windowAdmin = new()
            {
            };
            windowAdmin.Show();
        }

        private void Button_Review_Click(object sender, RoutedEventArgs e) => OpenReview();

        /// <summary>打开审核窗口（托盘菜单「审查」）</summary>
        internal void OpenReview()
        {
            if (!_permission.Can("review.view"))
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_ReviewNeedLogin"), LanguageService.Get("Cap_NoPermission"));
                return;
            }
            WindowReview windowReview = new()
            {
            };
            windowReview.Show();
        }
    }
}


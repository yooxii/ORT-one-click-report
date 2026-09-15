using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Admin.Views;
using ORT一键报告.Main.Views;
using ORT一键报告.Plans.Views;
using ORT一键报告.Reports.Views;
using ORT一键报告.Review.Views;
using ORT一键报告.Services;
using ORT一键报告.ViewModels;
using System;
using System.Collections.Generic;
using System.Windows;

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

        /// <summary>计划结束日期提醒的定时检查（每 6 小时一次）</summary>
        private readonly System.Windows.Threading.DispatcherTimer _mailTimer = new()
        {
            Interval = TimeSpan.FromHours(6)
        };

        public MainViewModel MainVM { get; set; }

        public MainWindow()
        {
            InitializeComponent();

            _auth = App.ServiceProvider.GetRequiredService<AuthService>();
            _permission = App.ServiceProvider.GetRequiredService<IPermissionService>();
            _reviewService = App.ServiceProvider.GetRequiredService<ReviewService>();
            _appSettings = App.ServiceProvider.GetRequiredService<AppSettingsService>();
            _mailNotifier = App.ServiceProvider.GetRequiredService<MailNotifier>();

            MainVM = App.ServiceProvider.GetRequiredService<MainViewModel>();
            DataContext = MainVM;
            MainVM.SubscribeLanguageChange();
            // 语言切换时刷新菜单文案与左下角用户身份信息
            LanguageService.LanguageChanged += () => Dispatcher.Invoke(UpdateUIByPermission);

            Loaded += (s, e) => Activate();
            // 启动时应用设置字体，并在设置变更时实时刷新所有已打开窗口
            Loaded += (s, e) => _appSettings.ApplyFont(this);
            _appSettings.SettingsChanged += () => Dispatcher.Invoke(() => _appSettings.ApplyFontToAll());

            _auth.AuthChanged += () => Dispatcher.Invoke(UpdateUIByPermission);
            Loaded += (s, e) => UpdateUIByPermission();
            Loaded += (s, e) => StartMailReminder();
            Closed += (s, e) => _mailTimer.Stop();
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
        /// 根据当前登录状态与角色刷新入口可用性，并关闭当前无权限访问的子窗口
        /// </summary>
        private void UpdateUIByPermission()
        {
            // 菜单项只显示登录/注销动作，当前用户身份信息统一在窗口左下角展示
            menu_account.Header = _auth.CurrentUser == null
                ? LanguageService.Get("Main_Login")
                : LanguageService.Get("Main_Logout");
            txt_user_identity.Text = string.Format(LanguageService.Get("Main_IdentityFormat"), _auth.CurrentDisplayName);
            btn_report.IsEnabled = _permission.Can("report.use");
            btn_admin.IsEnabled = _permission.Can("admin.manage");
            btn_review.IsEnabled = _permission.Can("review.view");
            btn_review.Content = _permission.Can("review.view")
                ? string.Format(LanguageService.Get("Main_ReviewCountFormat"), _reviewService.PendingCount())
                : LanguageService.Get("Main_Review");

            // 权限变化时关闭当前无权限访问的子窗口
            CloseUnauthorizedWindows();
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
            }
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

        private void MenuItem_Settings_Click(object sender, RoutedEventArgs e)
        {
            WindowAppSettings settingsWindow = new()
            {
                Owner = null
            };
            settingsWindow.Show();
        }

        private void MenuItem_Quit_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Button_YiJianBaoGao_Click(object sender, RoutedEventArgs e)
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
            WindowMainReport windowMainReport = new();
            windowMainReport.Show();
        }

        private void Button_Plans_Click(object sender, RoutedEventArgs e)
        {
            // 已打开的领退和计划窗口则聚焦，不重复打开
            foreach (Window w in Application.Current.Windows)
            {
                if (w is Plans.Views.WindowPlans existing)
                {
                    if (existing.WindowState == WindowState.Minimized)
                    {
                        existing.WindowState = WindowState.Normal;
                    }
                    existing.Activate();
                    return;
                }
            }
            Plans.Views.WindowPlans windowPlans = new()
            {
            };
            windowPlans.Show();
        }

        private void Button_Admin_Click(object sender, RoutedEventArgs e)
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

        private void Button_Review_Click(object sender, RoutedEventArgs e)
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


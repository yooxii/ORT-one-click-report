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

        public MainViewModel MainVM { get; set; }

        public MainWindow()
        {
            InitializeComponent();

            _auth = App.ServiceProvider.GetRequiredService<AuthService>();
            _permission = App.ServiceProvider.GetRequiredService<IPermissionService>();
            _reviewService = App.ServiceProvider.GetRequiredService<ReviewService>();
            _appSettings = App.ServiceProvider.GetRequiredService<AppSettingsService>();

            MainVM = App.ServiceProvider.GetRequiredService<MainViewModel>();
            DataContext = MainVM;
            MainVM.SubscribeLanguageChange();

            Loaded += (s, e) => Activate();
            // 启动时应用设置字体，并在设置变更时实时刷新
            Loaded += (s, e) => _appSettings.ApplyFont(this);
            _appSettings.SettingsChanged += () => Dispatcher.Invoke(() => _appSettings.ApplyFont(this));

            _auth.AuthChanged += () => Dispatcher.Invoke(UpdateUIByPermission);
            Loaded += (s, e) => UpdateUIByPermission();
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
            menu_account.Header = _auth.CurrentUser == null
                ? LanguageService.Get("Main_Login")
                : string.Format(LanguageService.Get("Main_LogoutFormat"), _auth.CurrentDisplayName);
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


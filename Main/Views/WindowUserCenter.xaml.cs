using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Models;
using ORT一键报告.Services;
using ORT一键报告.Services.Sim44;
using System;
using System.Linq;
using System.Windows;

namespace ORT一键报告.Main.Views
{
    /// <summary>
    /// 用户中心：管理当前登录用户自己的资料。
    /// 这里只保留「看一眼 / 改基本资料」，改密码、申请权限这类多步骤操作都收进了二级窗口
    /// （<see cref="WindowChangePassword"/>、<see cref="WindowRoleRequest"/>），避免一屏铺满。
    /// 入口：主界面菜单「用户 → 用户中心」。
    /// </summary>
    public partial class WindowUserCenter : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly AuthService _auth;
        private readonly ReviewService _review;
        private readonly Sim44Service _sim44;

        public WindowUserCenter()
        {
            InitializeComponent();
            _auth = App.ServiceProvider.GetRequiredService<AuthService>();
            _review = App.ServiceProvider.GetRequiredService<ReviewService>();
            _sim44 = App.ServiceProvider.GetRequiredService<Sim44Service>();
            Loaded += (_, _) => LoadUser();
        }

        /* ###############################  载入  ################################ */

        private void LoadUser()
        {
            User user = _auth.CurrentUser;
            if (user == null)
            {
                _ = MessageBox.Show(LanguageService.Get("UserCenter_Msg_NeedLogin"), LanguageService.Get("Cap_Info"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                Close();
                return;
            }
            txt_username.Text = user.Username;
            txt_roles.Text = _auth.CurrentRoles.Count == 0
                ? LanguageService.Get("Role_Guest")
                : string.Join("/", _auth.CurrentRoles.Select(r => LanguageService.Get("Role_" + r)));
            txt_createdAt.Text = user.CreatedAt.ToString("yyyy/M/d HH:mm");
            txt_displayName.Text = user.DisplayName ?? "";
            txt_email.Text = user.Email ?? "";

            UpdateCookieHint();
            UpdatePasswordState(user);
            UpdateRoleState(user);
        }

        private void UpdateCookieHint()
        {
            DateTime? expiry = _auth.GetCookieExpiry();
            txt_cookie.Text = expiry == null
                ? LanguageService.Get("UserCenter_NoCookie")
                : string.Format(LanguageService.Get("UserCenter_CookieExpiryFormat"), expiry.Value.ToString("yyyy/M/d HH:mm"));
        }

        /// <summary>密码入口的一句话状态：本程序密码有没有设 + MES 数据库密码还剩几天</summary>
        private void UpdatePasswordState(User user)
        {
            bool hasPassword = AuthService.HasPassword(user);
            string appState = hasPassword
                ? LanguageService.Get("UserCenter_PasswordSetState")
                : LanguageService.Get("UserCenter_NoPasswordHint");
            Sim44CredentialStatus status = _sim44.GetCredentialStatus();
            string dbState;
            if (!status.HasCredential)
            {
                dbState = LanguageService.Get("UserCenter_DbPasswordNotConfigured");
            }
            else if (status.DaysLeft is int days)
            {
                dbState = days < 0
                    ? string.Format(LanguageService.Get("Sim44_Status_ExpiredFormat"), -days)
                    : string.Format(LanguageService.Get("Sim44_Status_DaysLeftFormat"), days);
            }
            else
            {
                dbState = LanguageService.Get("Sim44_Status_ExpiryUnknown");
            }
            txt_passwordState.Text = appState + Environment.NewLine +
                                     string.Format(LanguageService.Get("UserCenter_DbPasswordStateFormat"), dbState);
        }

        /// <summary>权限入口的一句话状态：有没有待审批的申请 / 已经到最高权限</summary>
        private void UpdateRoleState(User user)
        {
            if (_auth.HasRole(UserRole.Administrator))
            {
                btn_roleRequest.IsEnabled = false;
                txt_roleState.Text = LanguageService.Get("UserCenter_RoleApplyAllGranted");
                return;
            }
            ReviewRequest pending = _review.GetPendingRoleRequest(user.Id);
            if (pending != null)
            {
                btn_roleRequest.IsEnabled = false;
                txt_roleState.Text = string.Format(LanguageService.Get("UserCenter_RoleApplyPendingFormat"),
                    pending.Action, pending.AssigneeName ?? "-");
                return;
            }
            btn_roleRequest.IsEnabled = true;
            txt_roleState.Text = LanguageService.Get("UserCenter_RoleApplyHint");
        }

        /* ###############################  事件函数  ################################ */

        private void Btn_Save_Click(object sender, RoutedEventArgs e)
        {
            if (_auth.CurrentUser == null)
            {
                return;
            }
            string error = _auth.SaveCurrentUserProfile(txt_displayName.Text, txt_email.Text);
            if (error != null)
            {
                _ = MessageBox.Show(this, error, LanguageService.Get("Cap_SaveFailed"));
                return;
            }
            _logger.Info($"用户中心：已保存 {_auth.CurrentUser.Username} 的资料");
            ToastService.Show(LanguageService.Get("UserCenter_Saved"), ToastType.Info);
            LoadUser();
        }

        private void Btn_ClearCookie_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show(this, LanguageService.Get("UserCenter_ClearCookieConfirm"), LanguageService.Get("Cap_Info"),
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            {
                return;
            }
            _auth.ClearLoginCookie();
            UpdateCookieHint();
            ToastService.Show(LanguageService.Get("UserCenter_CookieCleared"), ToastType.Info);
        }

        private void Btn_ChangePassword_Click(object sender, RoutedEventArgs e)
        {
            WindowChangePassword dialog = new() { Owner = this };
            dialog.ShowDialog();
            if (_auth.CurrentUser != null)
            {
                UpdatePasswordState(_auth.CurrentUser);
            }
        }

        private void Btn_RoleRequest_Click(object sender, RoutedEventArgs e)
        {
            WindowRoleRequest dialog = new() { Owner = this };
            dialog.ShowDialog();
            if (_auth.CurrentUser != null)
            {
                UpdateRoleState(_auth.CurrentUser);
            }
        }

        private void Btn_Close_Click(object sender, RoutedEventArgs e) => Close();
    }
}

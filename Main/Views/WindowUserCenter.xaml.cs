using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Models;
using ORT一键报告.Services;
using System;
using System.Linq;
using System.Windows;

namespace ORT一键报告.Main.Views
{
    /// <summary>
    /// 用户中心：管理当前登录用户自己的资料
    /// （显示名、邮箱、密码设置/修改，以及本机保存的登录信息）。
    /// 入口：主界面菜单「用户 → 用户中心」。
    /// </summary>
    public partial class WindowUserCenter : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly AuthService _auth;

        public WindowUserCenter()
        {
            InitializeComponent();
            _auth = App.ServiceProvider.GetRequiredService<AuthService>();
            Loaded += (s, e) => LoadUser();
        }

        /* ###############################  载入  ################################ */

        /// <summary>
        /// 载入当前登录用户的信息；未登录时提示并关闭窗口
        /// </summary>
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
            txt_currentPassword.Clear();
            txt_newPassword.Clear();
            txt_confirmPassword.Clear();

            // 没设置过密码时不需要"当前密码"，并且把提示改成"请设置密码"
            bool hasPassword = AuthService.HasPassword(user);
            Visibility currentVisibility = hasPassword ? Visibility.Visible : Visibility.Collapsed;
            lbl_currentPassword.Visibility = currentVisibility;
            txt_currentPassword.Visibility = currentVisibility;
            txt_passwordHint.Text = hasPassword
                ? LanguageService.Get("UserCenter_PasswordKeepHint")
                : LanguageService.Get("UserCenter_NoPasswordHint");

            UpdateCookieHint();
            if (hasPassword)
            {
                txt_displayName.Focus();
            }
            else
            {
                txt_newPassword.Focus();
            }
        }

        /// <summary>刷新"本机登录信息"（cookie 到期时间）</summary>
        private void UpdateCookieHint()
        {
            DateTime? expiry = _auth.GetCookieExpiry();
            txt_cookie.Text = expiry == null
                ? LanguageService.Get("UserCenter_NoCookie")
                : string.Format(LanguageService.Get("UserCenter_CookieExpiryFormat"), expiry.Value.ToString("yyyy/M/d HH:mm"));
        }

        /* ###############################  事件函数  ################################ */

        private void Btn_Save_Click(object sender, RoutedEventArgs e)
        {
            if (_auth.CurrentUser == null)
            {
                return;
            }
            bool hasPassword = AuthService.HasPassword(_auth.CurrentUser);
            string current = txt_currentPassword.Password;
            string newPassword = txt_newPassword.Password;
            string confirm = txt_confirmPassword.Password;

            // 1. 想改/设密码时先校验
            bool wantPasswordChange = !string.IsNullOrEmpty(newPassword) || !string.IsNullOrEmpty(confirm);
            if (hasPassword && wantPasswordChange && !_auth.VerifyCurrentUserPassword(current))
            {
                _ = MessageBox.Show(LanguageService.Get("Msg_CurrentPasswordWrong"), LanguageService.Get("Cap_Info"));
                txt_currentPassword.Focus();
                return;
            }
            if (wantPasswordChange)
            {
                if (newPassword != confirm)
                {
                    _ = MessageBox.Show(LanguageService.Get("Msg_PasswordMismatch"), LanguageService.Get("Cap_Info"));
                    txt_confirmPassword.Focus();
                    return;
                }
                if (newPassword.Length < 6)
                {
                    _ = MessageBox.Show(LanguageService.Get("Msg_PasswordTooShort"), LanguageService.Get("Cap_Info"));
                    txt_newPassword.Focus();
                    return;
                }
            }

            // 2. 基本资料（显示名 + 邮箱）
            string error = _auth.SaveCurrentUserProfile(txt_displayName.Text, txt_email.Text);
            if (error != null)
            {
                _ = MessageBox.Show(error, LanguageService.Get("Cap_SaveFailed"));
                return;
            }

            // 3. 密码
            if (wantPasswordChange)
            {
                error = _auth.SetCurrentUserPassword(newPassword);
                if (error != null)
                {
                    _ = MessageBox.Show(error, LanguageService.Get("Cap_SaveFailed"));
                    return;
                }
            }

            _logger.Info($"用户中心：已保存 {_auth.CurrentUser.Username} 的资料");
            ToastService.Show(LanguageService.Get("UserCenter_Saved"), ToastType.Info);
            LoadUser();
        }

        private void Btn_ClearCookie_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show(LanguageService.Get("UserCenter_ClearCookieConfirm"), LanguageService.Get("Cap_Info"),
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            {
                return;
            }
            _auth.ClearLoginCookie();
            UpdateCookieHint();
            ToastService.Show(LanguageService.Get("UserCenter_CookieCleared"), ToastType.Info);
        }

        private void Btn_Close_Click(object sender, RoutedEventArgs e) => Close();
    }
}

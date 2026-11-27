using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Models;
using ORT一键报告.Plans.Views;
using ORT一键报告.Services;
using ORT一键报告.Services.Sim44;
using System;
using System.Threading.Tasks;
using System.Windows;

namespace ORT一键报告.Main.Views
{
    /// <summary>
    /// 用户中心 →「修改密码」二级窗口：
    /// 1) 本程序账号密码（未设置过密码时这里是「设置密码」）；
    /// 2) MES 数据库账号密码（回线登记要写 SIM44 的那个账号，公司要求每 90 天变更一次）。
    /// </summary>
    public partial class WindowChangePassword : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly AuthService _auth;
        private readonly Sim44Service _sim44;
        private bool _busy;

        public WindowChangePassword()
        {
            InitializeComponent();
            _auth = App.ServiceProvider.GetRequiredService<AuthService>();
            _sim44 = App.ServiceProvider.GetRequiredService<Sim44Service>();
            Loaded += (_, _) => LoadAll();
        }

        private void LoadAll()
        {
            LoadAppPasswordSection();
            LoadDbPasswordSection();
        }

        /* ###############################  本程序密码  ################################ */

        private void LoadAppPasswordSection()
        {
            User user = _auth.CurrentUser;
            bool hasPassword = user != null && AuthService.HasPassword(user);
            Visibility currentVisibility = hasPassword ? Visibility.Visible : Visibility.Collapsed;
            lbl_currentPassword.Visibility = currentVisibility;
            txt_currentPassword.Visibility = currentVisibility;
            txt_passwordHint.Text = hasPassword
                ? LanguageService.Get("UserCenter_PasswordKeepHint")
                : LanguageService.Get("UserCenter_NoPasswordHint");
            txt_currentPassword.Clear();
            txt_newPassword.Clear();
            txt_confirmPassword.Clear();
        }

        private void Btn_SaveAppPassword_Click(object sender, RoutedEventArgs e)
        {
            User user = _auth.CurrentUser;
            if (user == null)
            {
                MessageBox.Show(this, LanguageService.Get("UserCenter_Msg_NeedLogin"), LanguageService.Get("Cap_Info"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            bool hasPassword = AuthService.HasPassword(user);
            string current = txt_currentPassword.Password;
            string newPassword = txt_newPassword.Password;
            string confirm = txt_confirmPassword.Password;

            if (string.IsNullOrEmpty(newPassword) && string.IsNullOrEmpty(confirm))
            {
                MessageBox.Show(this, LanguageService.Get("UserCenter_Msg_EmptyPassword"), LanguageService.Get("Cap_Info"),
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (hasPassword && !_auth.VerifyCurrentUserPassword(current))
            {
                MessageBox.Show(this, LanguageService.Get("Msg_CurrentPasswordWrong"), LanguageService.Get("Cap_Info"));
                txt_currentPassword.Focus();
                return;
            }
            if (newPassword != confirm)
            {
                MessageBox.Show(this, LanguageService.Get("Msg_PasswordMismatch"), LanguageService.Get("Cap_Info"));
                txt_confirmPassword.Focus();
                return;
            }
            if (newPassword.Length < 6)
            {
                MessageBox.Show(this, LanguageService.Get("Msg_PasswordTooShort"), LanguageService.Get("Cap_Info"));
                txt_newPassword.Focus();
                return;
            }
            string error = _auth.SetCurrentUserPassword(newPassword);
            if (error != null)
            {
                MessageBox.Show(this, error, LanguageService.Get("Cap_SaveFailed"));
                return;
            }
            _logger.Info($"用户中心：{user.Username} 修改了本程序密码");
            ToastService.Show(LanguageService.Get("UserCenter_Saved"), ToastType.Info);
            LoadAppPasswordSection();
        }

        /* ###############################  MES 数据库密码  ################################ */

        private void LoadDbPasswordSection()
        {
            Sim44CredentialStatus status = _sim44.GetCredentialStatus();
            if (!status.HasCredential)
            {
                grp_db.Visibility = Visibility.Collapsed;
                return;
            }
            grp_db.Visibility = Visibility.Visible;
            txt_dbUser.Text = status.User;
            if (status.DaysLeft is int days)
            {
                txt_dbExpiry.Text = days < 0
                    ? string.Format(LanguageService.Get("Sim44_Status_ExpiredFormat"), -days)
                    : string.Format(LanguageService.Get("Sim44_Status_DaysLeftFormat"), days);
            }
            else
            {
                txt_dbExpiry.Text = LanguageService.Get("Sim44_Status_ExpiryUnknown");
            }
            string lastChanged = status.LastChangedAt?.ToString("yyyy/M/d") ?? "-";
            txt_dbHint.Text = string.Format(LanguageService.Get("UserCenter_DbPasswordHintFormat"),
                Sim44CredentialStore.PasswordValidDays, lastChanged);
            txt_dbOldPassword.Clear();
            txt_dbNewPassword.Clear();
            txt_dbConfirmPassword.Clear();
        }

        private async void Btn_ChangeDbPassword_Click(object sender, RoutedEventArgs e)
        {
            if (_busy)
            {
                return;
            }
            string oldPassword = txt_dbOldPassword.Password;
            string newPassword = txt_dbNewPassword.Password;
            string confirm = txt_dbConfirmPassword.Password;
            if (string.IsNullOrEmpty(oldPassword))
            {
                MessageBox.Show(this, LanguageService.Get("UserCenter_Msg_CurrentDbPasswordRequired"),
                    LanguageService.Get("Cap_Info"));
                txt_dbOldPassword.Focus();
                return;
            }
            if (string.IsNullOrEmpty(newPassword))
            {
                MessageBox.Show(this, LanguageService.Get("UserCenter_Msg_EmptyPassword"), LanguageService.Get("Cap_Info"));
                txt_dbNewPassword.Focus();
                return;
            }
            if (newPassword != confirm)
            {
                MessageBox.Show(this, LanguageService.Get("Msg_PasswordMismatch"), LanguageService.Get("Cap_Info"));
                txt_dbConfirmPassword.Focus();
                return;
            }
            if (newPassword == oldPassword)
            {
                MessageBox.Show(this, LanguageService.Get("UserCenter_Msg_SamePassword"), LanguageService.Get("Cap_Info"));
                txt_dbNewPassword.Focus();
                return;
            }
            if (newPassword.Length < 6)
            {
                MessageBox.Show(this, LanguageService.Get("Msg_PasswordTooShort"), LanguageService.Get("Cap_Info"));
                txt_dbNewPassword.Focus();
                return;
            }
            if (MessageBox.Show(this, string.Format(LanguageService.Get("UserCenter_ConfirmChangeDbPassword"), newPassword),
                    LanguageService.Get("Cap_Info"), MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            {
                return;
            }

            _busy = true;
            btn_changeDbPassword.IsEnabled = false;
            txt_dbHint.Text = LanguageService.Get("Sim44_Status_Saving");
            try
            {
                // 改密需要一条有效连接：先试本机凭据，不行就让用户登进去
                if (!_sim44.IsLoggedIn && !await Task.Run(() => _sim44.TryAutoLogin(out _)))
                {
                    WindowSim44Login login = new(_sim44) { Owner = this };
                    if (login.ShowDialog() != true)
                    {
                        txt_dbHint.Text = LanguageService.Get("Sim44_Status_NotLoggedIn");
                        return;
                    }
                }
                (bool ok, string message) = await Task.Run(() => _sim44.ChangePassword(oldPassword, newPassword));
                if (ok)
                {
                    _logger.Info("用户中心：已修改 MES 数据库密码");
                    ToastService.Show(LanguageService.Get("UserCenter_DbPasswordChanged"), ToastType.Info);
                    LoadDbPasswordSection();
                }
                else
                {
                    txt_dbHint.Text = message;
                    MessageBox.Show(this, message, LanguageService.Get("Cap_SaveFailed"),
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            finally
            {
                _busy = false;
                btn_changeDbPassword.IsEnabled = true;
            }
        }

        private void Btn_Close_Click(object sender, RoutedEventArgs e) => Close();
    }
}

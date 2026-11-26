using Microsoft.Extensions.DependencyInjection;
using ORT一键报告.Models;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ORT一键报告.Main.Views
{
    /// <summary>
    /// WindowLogin.xaml 的交互逻辑。
    ///
    /// 两种登录方式：
    /// 1)「公司平台账号」（默认页签）：用公司 EIP 门户账号密码登录平台，平台验证通过即代表本人；
    ///    登录成功会返回真实姓名，这个名字就是本程序的用户名/显示名，进来以后的身份组按账号登记的身份走
    ///    （同部门账号首次登录自动登记，非本部门账号拒绝并提示找管理员）；
    /// 2)「本地账号」：保留原来的本地密码登录，给超级管理员（admin）和没开通平台账号的人兜底。
    /// 另有「游客登录」：不验证身份进入只读模式（沿用原有能力）。
    /// </summary>
    public partial class WindowLogin : Window
    {
        private readonly AuthService _auth;
        private readonly PortalAccountService _portal;

        /// <summary>检测到的有效登录 cookie（用户名/密码）</summary>
        private (string Username, string Password)? _cookie;

        /// <summary>cookie 对应的平台账号（不是平台账号则为 null，走本地登录继续）</summary>
        private PortalAccount _cookiePortalAccount;

        /// <summary>最近登录过的账号（快捷入口，选中即填账号）</summary>
        private List<PortalAccount> _recentAccounts = [];

        /// <summary>是否正在登录（避免重复点击）</summary>
        private bool _busy;

        /// <summary>
        /// 是否以游客身份进入（点了「游客登录」按钮）。
        /// 用于区分「游客进入」与「退出程序」：两者 DialogResult 都是 false，
        /// 调用方按该标志决定是打开主窗口还是关闭程序。
        /// </summary>
        public bool IsGuestMode { get; private set; }

        public WindowLogin()
        {
            InitializeComponent();
            _auth = App.ServiceProvider.GetRequiredService<AuthService>();
            _portal = App.ServiceProvider.GetRequiredService<PortalAccountService>();
            Loaded += (s, e) =>
            {
                LoadRecentAccounts();
                ShowCookieHint();
                txt_account.Focus();
            };
        }

        /* ###############################  载入  ################################ */

        /// <summary>最近登录过的平台账号填进下拉框（只列启用中的）</summary>
        private void LoadRecentAccounts()
        {
            try
            {
                _recentAccounts = _portal.RecentAccounts(8);
            }
            catch (Exception)
            {
                _recentAccounts = [];
            }
            foreach (PortalAccount account in _recentAccounts)
            {
                if (string.IsNullOrWhiteSpace(account.DisplayName))
                {
                    account.DisplayName = account.Account;
                }
            }
            cb_savedAccounts.ItemsSource = _recentAccounts;
            // 没有登记过的账号就把这一行藏起来（首次使用的人看不到空下拉框）
            Visibility visibility = _recentAccounts.Count == 0 ? Visibility.Collapsed : Visibility.Visible;
            lbl_savedAccounts.Visibility = visibility;
            cb_savedAccounts.Visibility = visibility;
        }

        /// <summary>
        /// 检测本地登录 cookie，有效时提示用户是否继续上一次登录。
        /// 平台账号的密码也留着（用于免密进入），但不是靠本地散列验，而是重新走一次平台登录。
        /// </summary>
        private void ShowCookieHint()
        {
            _cookie = _auth.LoadValidCookie();
            _cookiePortalAccount = null;
            if (_cookie == null)
            {
                return;
            }
            try
            {
                _cookiePortalAccount = _portal.FindByUserName(_cookie.Value.Username);
            }
            catch (Exception)
            {
                _cookiePortalAccount = null;
            }
            if (_cookiePortalAccount != null)
            {
                txt_account.Text = _cookiePortalAccount.Account;
            }
            else
            {
                txt_username.Text = _cookie.Value.Username;
            }
            DateTime? expiry = _auth.GetCookieExpiry();
            txt_cookieHint.Text = string.Format(LanguageService.Get("Login_CookieDetected"),
                _cookie.Value.Username, expiry?.ToString("yyyy/M/d") ?? "");
            panel_cookie.Visibility = Visibility.Visible;
        }

        /* ###############################  交互  ################################ */

        private void Tabs_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded)
            {
                return;
            }
            txt_error.Text = "";
            if (ReferenceEquals(tabs.SelectedItem, tab_local))
            {
                txt_username.Focus();
            }
            else
            {
                txt_account.Focus();
            }
        }

        /// <summary>选中「已保存的账号」：填账号 + 密码框留空（有保存的密码会提示可免密进入）</summary>
        private void Cb_SavedAccounts_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cb_savedAccounts.SelectedItem is not PortalAccount account)
            {
                return;
            }
            txt_account.Text = account.Account;
            txt_portalPassword.Clear();
            txt_error.Text = "";
            string stored = PortalAccountService.GetStoredPassword(account);
            if (!string.IsNullOrEmpty(stored))
            {
                txt_portalPassword.Password = stored;
                txt_error.Text = "";
            }
            txt_busy.Visibility = Visibility.Collapsed;
        }

        private void Txt_Password_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                Btn_Login_Click(sender, e);
            }
        }

        private void Btn_Login_Click(object sender, RoutedEventArgs e)
        {
            if (_busy)
            {
                return;
            }
            if (ReferenceEquals(tabs.SelectedItem, tab_local))
            {
                LoginLocal();
            }
            else
            {
                _ = LoginPortalAsync(txt_account.Text, txt_portalPassword.Password);
            }
        }

        /// <summary>
        /// 本地账号登录（超级管理员兜底）：沿用原来的本地密码散列校验；
        /// 密码可以留空——账号没设过密码时只凭用户名也能进（登录后提示去设置密码）。
        /// </summary>
        private void LoginLocal()
        {
            string username = txt_username.Text?.Trim();
            string password = txt_password.Password;
            if (string.IsNullOrEmpty(username))
            {
                txt_error.Text = LanguageService.Get("Login_EnterUsername");
                txt_username.Focus();
                return;
            }
            if (_auth.Login(username, password))
            {
                DialogResult = true;
            }
            else
            {
                txt_error.Text = string.IsNullOrEmpty(password)
                    ? LanguageService.Get("Login_PasswordRequiredHint")
                    : LanguageService.Get("Login_InvalidUserPass");
                txt_password.Clear();
                txt_password.Focus();
            }
        }

        /// <summary>
        /// 公司平台账号登录：验证平台 → 部门/身份判定 → 建登录态。失败按原因给不同提示。
        /// </summary>
        private async Task LoginPortalAsync(string account, string password)
        {
            account = (account ?? "").Trim();
            if (string.IsNullOrEmpty(account))
            {
                txt_error.Text = LanguageService.Get("WindowLogin_EnterAccount");
                txt_account.Focus();
                return;
            }
            if (string.IsNullOrEmpty(password))
            {
                txt_error.Text = LanguageService.Get("WindowLogin_EnterPassword");
                txt_portalPassword.Focus();
                return;
            }
            SetBusy(true, LanguageService.Get("WindowLogin_SigningIn"));
            try
            {
                PortalLoginOutcome outcome = await _portal.SignInAsync(account, password).ConfigureAwait(true);
                if (outcome.Success)
                {
                    DialogResult = true;
                    return;
                }
                txt_error.Text = string.IsNullOrWhiteSpace(outcome.Message)
                    ? LanguageService.Get("WindowLogin_LoginFailed")
                    : outcome.Message;
                if (outcome.Status == PortalLoginStatus.InvalidCredentials)
                {
                    txt_portalPassword.Clear();
                    txt_portalPassword.Focus();
                }
            }
            catch (Exception ex)
            {
                txt_error.Text = "登录时出错：" + ex.Message;
            }
            finally
            {
                SetBusy(false, null);
            }
        }

        /// <summary>「继续上次登录」：平台账号就再走一次平台登录，本地账号就按本地密码校验</summary>
        private void Btn_ContinueLogin_Click(object sender, RoutedEventArgs e)
        {
            if (_cookie == null || _busy)
            {
                return;
            }
            if (_cookiePortalAccount != null)
            {
                _ = ContinuePortalAsync();
                return;
            }
            if (_auth.Login(_cookie.Value.Username, _cookie.Value.Password))
            {
                DialogResult = true;
                return;
            }
            // 账号被禁用/密码已变更：清除失效 cookie，改为手动登录
            _auth.ClearLoginCookie();
            _cookie = null;
            panel_cookie.Visibility = Visibility.Collapsed;
            txt_error.Text = LanguageService.Get("Login_CookieExpired");
            txt_password.Focus();
        }

        private async Task ContinuePortalAsync()
        {
            SetBusy(true, LanguageService.Get("WindowLogin_SigningIn"));
            try
            {
                PortalLoginOutcome outcome = await _portal.SignInAsync(_cookiePortalAccount.Account,
                    _cookie.Value.Password).ConfigureAwait(true);
                if (outcome.Success)
                {
                    DialogResult = true;
                    return;
                }
                txt_error.Text = outcome.Message ?? LanguageService.Get("Login_CookieExpired");
                txt_account.Text = _cookiePortalAccount.Account;
            }
            catch (Exception ex)
            {
                txt_error.Text = "登录时出错：" + ex.Message;
            }
            finally
            {
                SetBusy(false, null);
            }
        }

        private void SetBusy(bool busy, string message)
        {
            _busy = busy;
            btn_login.IsEnabled = !busy;
            tabs.IsEnabled = !busy;
            txt_busy.Text = message ?? "";
            txt_busy.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
        }

        private void Btn_Cancel_Click(object sender, RoutedEventArgs e)
        {
            // 「退出」：不登录也不以游客进入，直接关闭程序（IsGuestMode 保持 false）
            DialogResult = false;
        }

        /// <summary>
        /// 「游客登录」：不验证账号，以游客身份（CurrentUser=null）进入主界面
        /// </summary>
        private void Btn_Guest_Click(object sender, RoutedEventArgs e)
        {
            IsGuestMode = true;
            DialogResult = false;
        }
    }
}

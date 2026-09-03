using Microsoft.Extensions.DependencyInjection;
using ORT一键报告.Services;
using System;
using System.Windows;
using System.Windows.Input;

namespace ORT一键报告.Main.Views
{
    /// <summary>
    /// WindowLogin.xaml 的交互逻辑
    /// </summary>
    public partial class WindowLogin : Window
    {
        private readonly AuthService _auth;

        /// <summary>
        /// 检测到的有效登录 cookie（用户名/密码）
        /// </summary>
        private (string Username, string Password)? _cookie;

        public WindowLogin()
        {
            InitializeComponent();
            _auth = App.ServiceProvider.GetRequiredService<AuthService>();
            Loaded += (s, e) =>
            {
                ShowCookieHint();
                txt_username.Focus();
            };
        }

        /// <summary>
        /// 检测本地登录 cookie，有效时提示用户是否继续上一次登录
        /// </summary>
        private void ShowCookieHint()
        {
            _cookie = _auth.LoadValidCookie();
            if (_cookie == null)
            {
                return;
            }
            txt_username.Text = _cookie.Value.Username;
            DateTime? expiry = _auth.GetCookieExpiry();
            txt_cookieHint.Text = string.Format(LanguageService.Get("Login_CookieDetected"), _cookie.Value.Username, expiry?.ToString("yyyy/M/d") ?? "");
            panel_cookie.Visibility = Visibility.Visible;
        }

        private void Btn_ContinueLogin_Click(object sender, RoutedEventArgs e)
        {
            if (_cookie == null)
            {
                return;
            }
            if (_auth.Login(_cookie.Value.Username, _cookie.Value.Password))
            {
                DialogResult = true;
            }
            else
            {
                // 账号被禁用/密码已变更：清除失效 cookie，改为手动登录
                _auth.ClearLoginCookie();
                _cookie = null;
                panel_cookie.Visibility = Visibility.Collapsed;
                txt_error.Text = LanguageService.Get("Login_CookieExpired");
                txt_password.Focus();
            }
        }

        private void Btn_Login_Click(object sender, RoutedEventArgs e)
        {
            string username = txt_username.Text?.Trim();
            string password = txt_password.Password;
            if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
            {
                txt_error.Text = LanguageService.Get("Login_EnterUserPass");
                return;
            }
            if (_auth.Login(username, password))
            {
                DialogResult = true;
            }
            else
            {
                txt_error.Text = LanguageService.Get("Login_InvalidUserPass");
                txt_password.Clear();
                txt_password.Focus();
            }
        }

        private void Txt_Password_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                Btn_Login_Click(sender, e);
            }
        }

        private void Btn_Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}

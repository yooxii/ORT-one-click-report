using Microsoft.Extensions.DependencyInjection;
using ORT一键报告.Services;
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

        public WindowLogin()
        {
            InitializeComponent();
            _auth = App.ServiceProvider.GetRequiredService<AuthService>();
            Loaded += (s, e) => txt_username.Focus();
        }

        private void Btn_Login_Click(object sender, RoutedEventArgs e)
        {
            string username = txt_username.Text?.Trim();
            string password = txt_password.Password;
            if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
            {
                txt_error.Text = "请输入用户名和密码";
                return;
            }
            if (_auth.Login(username, password))
            {
                DialogResult = true;
            }
            else
            {
                txt_error.Text = "用户名或密码错误，或账号已禁用";
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

using NLog;
using ORT一键报告.Services;
using ORT一键报告.Services.Sim44;
using System;
using System.Threading.Tasks;
using System.Windows;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// MES 数据库登录框：用户没在本机存过密码（或存的不对）时弹出。
    /// 勾「记住密码」会用 DPAPI 加密保存到当前 Windows 用户目录，下次自动登录。
    /// </summary>
    public partial class WindowSim44Login : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly Sim44Service _service;
        private bool _busy;

        public WindowSim44Login(Sim44Service service)
        {
            InitializeComponent();
            _service = service ?? throw new ArgumentNullException(nameof(service));
            Loaded += (_, _) =>
            {
                Sim44CredentialStatus status = _service.GetCredentialStatus();
                txt_user.Text = string.IsNullOrWhiteSpace(status.User) ? Environment.UserName : status.User;
                if (!string.IsNullOrWhiteSpace(txt_user.Text))
                {
                    pwd.Focus();
                }
                else
                {
                    txt_user.Focus();
                }
            };
        }

        private async void Btn_Login_Click(object sender, RoutedEventArgs e)
        {
            if (_busy)
            {
                return;
            }
            string user = txt_user.Text.Trim();
            string password = pwd.Password;
            // 记住密码的勾选状态必须在 UI 线程读出来：下面丢到后台线程做登录，
            // 后台线程碰 WPF 控件会抛「调用线程无法访问此对象」。
            bool remember = chk_remember.IsChecked == true;
            if (string.IsNullOrWhiteSpace(user))
            {
                ShowMessage(LanguageService.Get("Sim44_Msg_UserRequired"));
                txt_user.Focus();
                return;
            }
            if (string.IsNullOrEmpty(password))
            {
                ShowMessage(LanguageService.Get("Sim44_Msg_PasswordRequired"));
                pwd.Focus();
                return;
            }

            _busy = true;
            btn_ok.IsEnabled = false;
            txt_msg.Text = LanguageService.Get("Sim44_Status_LoggingIn");
            try
            {
                (Sim44Service.LoginOutcome outcome, string actualUser, string message) =
                    await Task.Run(() => _service.Login(user, password, remember));
                switch (outcome)
                {
                    case Sim44Service.LoginOutcome.Success:
                        _logger.Info($"MES 数据库登录成功：{actualUser}"
                                     + (string.IsNullOrWhiteSpace(message) ? "" : "｜警告：" + message));
                        DialogResult = true;
                        return;
                    case Sim44Service.LoginOutcome.BadCredential:
                        ShowMessage(LanguageService.Get("Sim44_Msg_BadCredential"));
                        pwd.SelectAll();
                        pwd.Focus();
                        break;
                    case Sim44Service.LoginOutcome.PasswordExpired:
                        ShowMessage(LanguageService.Get("Sim44_Msg_PasswordExpired"));
                        break;
                    case Sim44Service.LoginOutcome.NoTable:
                        ShowMessage(LanguageService.Get("Sim44_Msg_NoTable") + Environment.NewLine + message);
                        break;
                    case Sim44Service.LoginOutcome.Network:
                        ShowMessage(LanguageService.Get("Sim44_Msg_Network") + Environment.NewLine + message);
                        break;
                    case Sim44Service.LoginOutcome.Environment:
                        ShowMessage(LanguageService.Get("Sim44_Msg_Environment") + Environment.NewLine + message);
                        break;
                    default:
                        ShowMessage(message);
                        break;
                }
            }
            finally
            {
                _busy = false;
                btn_ok.IsEnabled = true;
            }
        }

        private void ShowMessage(string text)
        {
            txt_msg.Text = text;
        }

        private async void Btn_SelfTest_Click(object sender, RoutedEventArgs e)
        {
            btn_selftest.IsEnabled = false;
            try
            {
                txt_msg.Text = LanguageService.Get("Sim44_Status_SelfTesting");
                (bool ok, string message) = await Task.Run(() => _service.SelfTest());
                string caption = LanguageService.Get(ok ? "Sim44_SelfTest_Ok" : "Sim44_SelfTest_Fail");
                txt_msg.Text = message;
                MessageBox.Show(this, message, caption, MessageBoxButton.OK,
                    ok ? MessageBoxImage.Information : MessageBoxImage.Warning);
            }
            finally
            {
                btn_selftest.IsEnabled = true;
            }
        }

        /// <summary>改连接设置（数据源/表名/桥接路径）：登录成功却报表不存在时从这里改</summary>
        private void Btn_Settings_Click(object sender, RoutedEventArgs e)
        {
            WindowSim44Settings settings = new() { Owner = this };
            settings.ShowDialog();
            // 设置改过之后提示重新登录，避免用旧设置继续试
            ShowMessage(LanguageService.Get("Sim44_Settings_ReloginHint"));
        }

        private void Btn_Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;
    }
}

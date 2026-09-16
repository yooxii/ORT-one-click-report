using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ORT一键报告
{
    /// <summary>
    /// PopupWindow.xaml 的交互逻辑
    /// </summary>
    public partial class PopupWindow : Window
    {
        public string Message { get; set; } = "这是一个弹出窗口";
        public string Result { get; set; }
        private List<ButtonConfig> _buttons;

        private struct ButtonConfig
        {
            public string Text;
            public string Result;

            public ButtonConfig(string text, string result)
            {
                Text = text;
                Result = result;
            }
        }

        public PopupWindow()
        {
            InitializeComponent();
            DataContext = this;
            _buttons = new List<ButtonConfig>();
            // 没配按钮 = 当"处理中/请稍候"窗口用：显示不确定进度条
            if (_buttons.Count == 0)
            {
                BusyBar.Visibility = Visibility.Visible;
            }
            Loaded += (s, e) => _shownAt = DateTime.Now;
            Closing += OnBusyClosingDelay;
        }

        /// <summary>等待窗口至少显示这么久：操作很快时也有一眼反馈，不会一闪而过</summary>
        private const int BusyMinVisibleMs = 400;

        private DateTime _shownAt = DateTime.Now;
        private bool _delayedClose;

        /// <summary>等待窗口（无按钮）显示不足 BusyMinVisibleMs 时延后关闭</summary>
        private void OnBusyClosingDelay(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (_buttons.Count > 0 || _delayedClose)
            {
                return;
            }
            double elapsed = (DateTime.Now - _shownAt).TotalMilliseconds;
            if (elapsed >= BusyMinVisibleMs)
            {
                return;
            }
            e.Cancel = true;
            _delayedClose = true;
            System.Windows.Threading.DispatcherTimer timer = new()
            {
                Interval = TimeSpan.FromMilliseconds(BusyMinVisibleMs - elapsed)
            };
            timer.Tick += (s, args) =>
            {
                timer.Stop();
                Close();
            };
            timer.Start();
        }

        /// <summary>
        /// 显示"正在处理"等待窗口（带不确定进度条）。返回窗口实例，调用方处理完自行 Close()。
        /// </summary>
        public static PopupWindow ShowBusy(string message, Window owner = null)
        {
            PopupWindow window = new()
            {
                Title = LanguageService.Get("Title_Processing"),
                Message = message
            };
            window.Owner = owner ?? FindActiveWindow();
            window.BusyBar.Visibility = Visibility.Visible;
            window.IconBadge.Visibility = Visibility.Collapsed;
            window.ButtonPanel.Visibility = Visibility.Collapsed;
            window.Show();
            return window;
        }

        /// <summary>取当前活动（或主）窗口，作为等待窗口的属主，保证居中显示</summary>
        private static Window FindActiveWindow()
        {
            if (Application.Current == null)
            {
                return null;
            }
            foreach (Window w in Application.Current.Windows)
            {
                if (w != null && w.IsVisible && w.IsActive)
                {
                    return w;
                }
            }
            return Application.Current.MainWindow is { IsVisible: true } main ? main : null;
        }

        public static string Show(string message, string title, MessageBoxImage icon, params (string Text, string Result)[] buttons)
        {
            var window = new PopupWindow();
            window.Configure(message, title, icon, buttons);

            if (Application.Current != null)
            {
                if (Application.Current.MainWindow != null && Application.Current.MainWindow.IsVisible)
                {
                    window.Owner = Application.Current.MainWindow;
                }
                else
                {
                    foreach (Window w in Application.Current.Windows)
                    {
                        if (w != null && w.IsVisible && w.IsActive)
                        {
                            window.Owner = w;
                            break;
                        }
                    }
                }
            }

            bool? dialogResult = window.ShowDialog();
            return dialogResult.HasValue && dialogResult.Value
                ? window.Result
                : string.Empty;
        }

        // 重载：无图标
        public static string Show(string message, string title, params (string Text, string Result)[] buttons)
        {
            return Show(message, title, MessageBoxImage.None, buttons);
        }

        public void Configure(string message, string title, MessageBoxImage icon, params (string Text, string Result)[] buttons)
        {
            if (buttons == null || buttons.Length == 0)
            {
                throw new ArgumentException("至少需要一个按钮");
            }

            Message = message;
            Title = title;

            _buttons.Clear();
            foreach ((string Text, string Result) btn in buttons)
            {
                _buttons.Add(new ButtonConfig(btn.Text, btn.Result));
            }

            BusyBar.Visibility = Visibility.Collapsed;
            SetIcon(icon);
            CreateButtons();
        }

        /// <summary>
        /// 图标、颜色都走主题语义键（错误红 / 警告黄 / 成功绿），未设置图标时隐藏图标徽标
        /// </summary>
        public void SetIcon(MessageBoxImage icon)
        {
            string glyph;
            string colorKey;
            string badgeKey;
            switch (icon)
            {
                case MessageBoxImage.Error:
                    glyph = "\xE783";
                    colorKey = "StatusErrorBrush";
                    badgeKey = "StatusErrorBgBrush";
                    break;
                case MessageBoxImage.Question:
                case MessageBoxImage.Warning:
                    glyph = icon == MessageBoxImage.Question ? "\xE11D" : "\xE7BA";
                    colorKey = "StatusWarnBrush";
                    badgeKey = "StatusWarnBgBrush";
                    break;
                case MessageBoxImage.Information:
                    glyph = "\xE946";
                    colorKey = "StatusOkBrush";
                    badgeKey = "StatusOkBgBrush";
                    break;
                default:
                    glyph = "";
                    colorKey = "StatusOkBrush";
                    badgeKey = "StatusOkBgBrush";
                    break;
            }
            IconText.Text = glyph;
            IconBadge.Visibility = string.IsNullOrEmpty(glyph) ? Visibility.Collapsed : Visibility.Visible;
            IconText.SetResourceReference(TextBlock.ForegroundProperty, colorKey);
            IconBadge.SetResourceReference(Border.BackgroundProperty, badgeKey);
        }

        private void CreateButtons()
        {
            ButtonPanel.Children.Clear();

            // 从右到左添加（符合 Windows 习惯）
            for (int i = _buttons.Count - 1; i >= 0; i--)
            {
                var config = _buttons[i];
                var btn = new Button
                {
                    Content = config.Text,
                    Width = Math.Max(80, config.Text.Length * 12),
                    Height = 32,
                    Margin = new Thickness(5, 0, 0, 0),
                    Padding = new Thickness(10, 0, 10, 0)
                };

                // 捕获当前值（避免闭包陷阱）
                string currentResult = config.Result;
                btn.Click += (s, e) =>
                {
                    Result = currentResult;
                    DialogResult = true;
                    Close();
                };

                ButtonPanel.Children.Add(btn);
            }
            ButtonPanel.Visibility = _buttons.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
            BusyBar.Visibility = _buttons.Count > 0 ? Visibility.Collapsed : Visibility.Visible;
        }
    }
}

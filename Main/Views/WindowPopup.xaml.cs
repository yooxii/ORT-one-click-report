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

            SetIcon(icon);
            CreateButtons();
        }

        public void SetIcon(MessageBoxImage icon)
        {
            string glyph, color;
            switch (icon)
            {
                case MessageBoxImage.Error:
                    glyph = "\xE783";
                    color = "#D0021B";
                    break;
                case MessageBoxImage.Question:
                    glyph = "\xE11D";
                    color = "#4A90E2";
                    break;
                case MessageBoxImage.Warning:
                    glyph = "\xE7BA";
                    color = "#4A90E2";
                    break;
                case MessageBoxImage.Information:
                    glyph = "\xE946";
                    color = "#50C878";
                    break;
                case MessageBoxImage.None:
                    glyph = "";
                    color = "#50C878";
                    break;
                default:
                    glyph = "\xE946";
                    color = "#50C878";
                    break;
            }
            IconText.Text = glyph;
            IconText.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(color));
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
        }
    }
}

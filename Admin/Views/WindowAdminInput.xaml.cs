using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace ORT一键报告.Admin.Views
{
    /// <summary>
    /// 通用输入对话框：按字段定义动态生成输入行（TextBox/PasswordBox），返回输入值列表。
    /// </summary>
    public partial class WindowAdminInput : Window
    {
        private readonly List<Control> _inputs = [];

        /// <summary>
        /// 用户输入的值（与字段定义顺序一致）
        /// </summary>
        public List<string> Values { get; } = [];

        /// <summary>
        /// 构造输入对话框
        /// </summary>
        /// <param name="title">窗口标题</param>
        /// <param name="fields">(标签, 初始值, 是否密码框)</param>
        public WindowAdminInput(string title, params (string Label, string Initial, bool IsPassword)[] fields)
        {
            InitializeComponent();
            Title = title;

            foreach ((string label, string initial, bool isPassword) in fields)
            {
                Grid row = new();
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(90) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                row.Margin = new Thickness(0, 3, 0, 3);

                row.Children.Add(new Label { Content = label, VerticalContentAlignment = VerticalAlignment.Center });
                Control input;
                if (isPassword)
                {
                    PasswordBox pb = new() { VerticalContentAlignment = VerticalAlignment.Center };
                    if (!string.IsNullOrEmpty(initial))
                    {
                        pb.Password = initial;
                    }
                    input = pb;
                }
                else
                {
                    input = new TextBox { Text = initial ?? "", VerticalContentAlignment = VerticalAlignment.Center };
                }
                Grid.SetColumn(input, 1);
                row.Children.Add(input);

                panel_fields.Children.Add(row);
                _inputs.Add(input);
            }
            Loaded += (s, e) => _inputs[0]?.Focus();
        }

        private void Btn_OK_Click(object sender, RoutedEventArgs e)
        {
            Values.Clear();
            foreach (Control input in _inputs)
            {
                Values.Add(input is PasswordBox pb ? pb.Password : ((TextBox)input).Text?.Trim());
            }
            DialogResult = true;
        }

        private void Btn_Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}

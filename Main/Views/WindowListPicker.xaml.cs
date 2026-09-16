using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace ORT一键报告.Main.Views
{
    /// <summary>
    /// 通用单选对话框：给一串"显示名 → 值"，让用户选一个。
    /// 目前用于"报告文件夹里有多个报告文件时，选择要查看的那一份"。
    /// </summary>
    public partial class WindowListPicker : Window
    {
        /// <summary>列表项：Display 显示文本，Value 实际值</summary>
        public sealed class ListItem
        {
            public string Display { get; set; }
            public string Value { get; set; }
        }

        /// <summary>选中项的值（取消时为 null）</summary>
        public string SelectedValue { get; private set; }

        private WindowListPicker()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 弹出单选对话框；返回选中项的值，取消返回 null
        /// </summary>
        public static string Pick(Window owner, string title, string hint, IEnumerable<ListItem> items)
        {
            List<ListItem> list = (items ?? []).ToList();
            if (list.Count == 0)
            {
                return null;
            }
            if (list.Count == 1)
            {
                return list[0].Value;   // 只有一个就不用弹窗了
            }
            WindowListPicker dialog = new()
            {
                Owner = owner,
                Title = title
            };
            dialog.txt_hint.Text = hint ?? "";
            dialog.lb_items.ItemsSource = list;
            dialog.lb_items.SelectedIndex = 0;
            return dialog.ShowDialog() == true ? dialog.SelectedValue : null;
        }

        private void Lb_Items_MouseDoubleClick(object sender, MouseButtonEventArgs e) => Confirm();

        private void Btn_Ok_Click(object sender, RoutedEventArgs e) => Confirm();

        private void Confirm()
        {
            if (lb_items.SelectedItem is ListItem item)
            {
                SelectedValue = item.Value;
                DialogResult = true;
            }
        }
    }
}

using ORT一键报告.Models;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace ORT一键报告.Admin.Views
{
    /// <summary>
    /// 测试项目新增/编辑对话框：负责人支持从已有技术员中多选，也可直接输入（多个以"/"分隔）。
    /// </summary>
    public partial class WindowTestItemEdit : Window
    {
        private readonly List<CheckBox> _ownerBoxes = [];

        /// <summary>
        /// 抑制"文本框 ↔ 复选框"互相更新时的递归
        /// </summary>
        private bool _syncing;

        public string ItemName => txt_name.Text?.Trim();
        public string Period => txt_period.Text?.Trim();
        public string Category => cmb_category.Text?.Trim();
        public string Owner => txt_owner.Text?.Trim();
        public string Remark => txt_remark.Text?.Trim();

        /// <param name="title">窗口标题（新增/编辑）</param>
        /// <param name="item">待编辑的测试项目（新增时传空对象）</param>
        /// <param name="technicians">可选的技术员列表（已有技术员）</param>
        public WindowTestItemEdit(string title, TestItemCatalog item, IEnumerable<UserView> technicians)
        {
            InitializeComponent();
            Title = title;
            txt_name.Text = item?.Name ?? "";
            txt_period.Text = item?.Period ?? "";
            cmb_category.ItemsSource = TestCategories.Known.ToList();
            cmb_category.Text = string.IsNullOrWhiteSpace(item?.Category)
                ? TestCategories.Classify(item?.Name)
                : item.Category;
            txt_remark.Text = item?.Remark ?? "";
            txt_owner.Text = item?.Owner ?? "";
            BuildTechnicianList(technicians ?? []);
            SyncBoxesFromText();
        }

        /* ###############################  功能函数  ################################ */

        /// <summary>
        /// 生成已有技术员的复选框列表（勾选即加入负责人）
        /// </summary>
        private void BuildTechnicianList(IEnumerable<UserView> technicians)
        {
            foreach (UserView tech in technicians)
            {
                string name = string.IsNullOrWhiteSpace(tech.DisplayName) ? tech.Username : tech.DisplayName;
                string label = string.Equals(name, tech.Username, StringComparison.CurrentCultureIgnoreCase)
                    ? name
                    : $"{name} ({tech.Username})";
                CheckBox box = new()
                {
                    Content = label,
                    Tag = name,
                    Margin = new Thickness(2, 1, 2, 1)
                };
                box.Checked += OwnerBox_Changed;
                box.Unchecked += OwnerBox_Changed;
                _ownerBoxes.Add(box);
                panel_technicians.Children.Add(box);
            }
        }

        /// <summary>
        /// 复选框 → 负责人文本（保留文本中不在技术员列表内的自定义负责人）
        /// </summary>
        private void OwnerBox_Changed(object sender, RoutedEventArgs e)
        {
            if (_syncing)
            {
                return;
            }
            _syncing = true;
            try
            {
                List<string> known = _ownerBoxes.Select(b => (string)b.Tag).ToList();
                List<string> checkedNames = _ownerBoxes
                    .Where(b => b.IsChecked == true)
                    .Select(b => (string)b.Tag)
                    .ToList();
                List<string> custom = AdminService.SplitOwners(txt_owner.Text)
                    .Where(t => !known.Contains(t, StringComparer.CurrentCultureIgnoreCase))
                    .ToList();
                txt_owner.Text = string.Join("/", checkedNames.Concat(custom));
            }
            finally
            {
                _syncing = false;
            }
        }

        /// <summary>
        /// 负责人文本 → 复选框状态
        /// </summary>
        private void SyncBoxesFromText()
        {
            List<string> tokens = AdminService.SplitOwners(txt_owner.Text);
            _syncing = true;
            try
            {
                foreach (CheckBox box in _ownerBoxes)
                {
                    box.IsChecked = tokens.Contains((string)box.Tag, StringComparer.CurrentCultureIgnoreCase);
                }
            }
            finally
            {
                _syncing = false;
            }
        }

        /* ###############################  事件函数  ################################ */

        private void Txt_Owner_TextChanged(object sender, TextChangedEventArgs e) => SyncBoxesFromText();

        private void Btn_OK_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_name.Text))
            {
                _ = MessageBox.Show(LanguageService.Get("Msg_TestItemNameRequired"), LanguageService.Get("Cap_Info"));
                txt_name.Focus();
                return;
            }
            DialogResult = true;
        }

        private void Btn_Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;
    }
}

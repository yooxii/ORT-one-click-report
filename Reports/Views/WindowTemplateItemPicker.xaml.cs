using ORT一键报告.Main.Views;
using ORT一键报告.Models;
using ORT一键报告.Services;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ORT一键报告.Reports.Views
{
    /// <summary>
    /// 从"测试项目表"（test_items_catalog）里挑选测试项：支持搜索、多选、双击确定。
    /// 报告模板工具的「添加测试项」用它替代原来的手工输入。
    /// </summary>
    public partial class WindowTemplateItemPicker : Window
    {
        private readonly List<TestItemCatalog> _all;
        private readonly ObservableCollection<TestItemCatalog> _view = [];

        /// <summary>用户选中的测试项（取消时为空列表）</summary>
        public List<TestItemCatalog> Selected { get; private set; } = [];

        private WindowTemplateItemPicker(IEnumerable<TestItemCatalog> items)
        {
            InitializeComponent();
            _all = (items ?? []).ToList();
            dg_items.ItemsSource = _view;
            ApplyFilter();
            Loaded += (s, e) => txt_search.Focus();
        }

        /// <summary>
        /// 弹出选择器。alreadyAddedNames 里的测试项不再列出（避免重复添加）
        /// </summary>
        public static List<TestItemCatalog> Pick(Window owner, IEnumerable<TestItemCatalog> items, IEnumerable<string> alreadyAddedNames)
        {
            HashSet<string> added = new((alreadyAddedNames ?? []).Select(PlanIndexService.NameKey));
            List<TestItemCatalog> candidates = (items ?? [])
                .Where(i => !string.IsNullOrWhiteSpace(i.Name) && !added.Contains(PlanIndexService.NameKey(i.Name)))
                .ToList();
            WindowTemplateItemPicker dialog = new(candidates) { Owner = owner };
            return dialog.ShowDialog() == true ? dialog.Selected : [];
        }

        private void ApplyFilter()
        {
            string keyword = PlanIndexService.NameKey(txt_search.Text ?? "");
            _view.Clear();
            foreach (TestItemCatalog item in _all)
            {
                if (keyword.Length == 0 || PlanIndexService.NameKey(item.Name ?? "").Contains(keyword))
                {
                    _view.Add(item);
                }
            }
            txt_hint.Text = string.Format(LanguageService.Get("ReportTemplate_PickItemCount"), _all.Count, _view.Count);
        }

        private void Txt_Search_TextChanged(object sender, TextChangedEventArgs e) => ApplyFilter();

        private void Dg_Items_MouseDoubleClick(object sender, MouseButtonEventArgs e) => Confirm();

        private void Btn_Ok_Click(object sender, RoutedEventArgs e) => Confirm();

        private void Confirm()
        {
            if (dg_items.SelectedItems.Count == 0)
            {
                ToastService.Show(LanguageService.Get("ReportTemplate_Msg_PickItemRequired"), ToastType.Warning);
                return;
            }
            Selected = dg_items.SelectedItems.Cast<TestItemCatalog>().ToList();
            DialogResult = true;
        }
    }
}

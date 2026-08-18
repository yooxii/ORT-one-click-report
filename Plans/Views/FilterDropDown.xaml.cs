using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// Excel 风格的可复用筛选下拉控件：按钮 + 下拉面板（搜索框 / 全选 / 复选列表）。
    /// SelectedValues 为空集合表示"全部"（不过滤）。
    /// </summary>
    public partial class FilterDropDown : UserControl
    {
        /// <summary>
        /// 下拉选项项
        /// </summary>
        public class OptionItem : INotifyPropertyChanged
        {
            public string Name { get; }

            private bool _isChecked = true;
            public bool IsChecked
            {
                get => _isChecked;
                set
                {
                    if (_isChecked != value)
                    {
                        _isChecked = value;
                        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsChecked)));
                    }
                }
            }

            public OptionItem(string name) => Name = name;

            public event PropertyChangedEventHandler PropertyChanged;
        }

        /// <summary>
        /// 选项列表（由外部绑定提供）
        /// </summary>
        public static readonly DependencyProperty OptionsProperty =
            DependencyProperty.Register(nameof(Options), typeof(IEnumerable<string>), typeof(FilterDropDown),
                new PropertyMetadata(null, (d, e) => ((FilterDropDown)d).RebuildItems()));
        public IEnumerable<string> Options
        {
            get => (IEnumerable<string>)GetValue(OptionsProperty);
            set => SetValue(OptionsProperty, value);
        }

        /// <summary>
        /// 当前选中的值集合（双向绑定；空集合表示"全部"）。控件只修改集合内容，不替换集合实例。
        /// </summary>
        public static readonly DependencyProperty SelectedValuesProperty =
            DependencyProperty.Register(nameof(SelectedValues), typeof(IList<string>), typeof(FilterDropDown),
                new PropertyMetadata(null, (d, e) => ((FilterDropDown)d).ApplySelectedValues()));
        public IList<string> SelectedValues
        {
            get => (IList<string>)GetValue(SelectedValuesProperty);
            set => SetValue(SelectedValuesProperty, value);
        }

        /// <summary>
        /// 按钮显示的摘要文本
        /// </summary>
        public static readonly DependencyProperty SummaryTextProperty =
            DependencyProperty.Register(nameof(SummaryText), typeof(string), typeof(FilterDropDown),
                new PropertyMetadata("(全部)"));
        public string SummaryText
        {
            get => (string)GetValue(SummaryTextProperty);
            private set => SetValue(SummaryTextProperty, value);
        }

        /// <summary>
        /// 选择发生变化时触发（用于外部刷新过滤）
        /// </summary>
        public event EventHandler SelectionChanged;

        private readonly ObservableCollection<OptionItem> _visibleItems = [];
        private readonly List<OptionItem> _allItems = [];
        private bool _updating;

        public FilterDropDown()
        {
            InitializeComponent();
            PART_List.ItemsSource = _visibleItems;
            PART_Search.TextChanged += (s, e) => ApplySearch();
        }

        /* ###############################  内部逻辑  ################################ */

        /// <summary>
        /// 选项变化时重建列表（默认全选，即不过滤）
        /// </summary>
        private void RebuildItems()
        {
            _updating = true;
            _allItems.Clear();
            foreach (string option in Options ?? Enumerable.Empty<string>())
            {
                _allItems.Add(new OptionItem(option));
            }
            ApplySelectedValuesInternal();
            ApplySearch();
            _updating = false;
            UpdateSummary();
        }

        /// <summary>
        /// 外部 SelectedValues 变化时同步勾选状态
        /// </summary>
        private void ApplySelectedValues()
        {
            if (_updating)
            {
                return;
            }
            _updating = true;
            ApplySelectedValuesInternal();
            _updating = false;
            UpdateSummary();
        }

        private void ApplySelectedValuesInternal()
        {
            IList<string> selected = SelectedValues;
            // 空集合或null表示"全部"
            bool all = selected == null || selected.Count == 0;
            foreach (OptionItem item in _allItems)
            {
                item.IsChecked = all || selected.Contains(item.Name);
            }
            SyncSelectAllState();
        }

        /// <summary>
        /// 搜索框过滤可见选项（不改变勾选状态）
        /// </summary>
        private void ApplySearch()
        {
            string keyword = PART_Search?.Text?.Trim();
            _visibleItems.Clear();
            foreach (OptionItem item in _allItems)
            {
                if (string.IsNullOrEmpty(keyword)
                    || item.Name.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    _visibleItems.Add(item);
                }
            }
        }

        private void Item_CheckChanged(object sender, RoutedEventArgs e)
        {
            if (_updating)
            {
                return;
            }
            SyncToSelectedValues();
        }

        private void SelectAll_Changed(object sender, RoutedEventArgs e)
        {
            if (_updating || PART_SelectAll == null)
            {
                return;
            }
            _updating = true;
            bool isChecked = PART_SelectAll.IsChecked == true;
            foreach (OptionItem item in _allItems)
            {
                item.IsChecked = isChecked;
            }
            _updating = false;
            SyncToSelectedValues();
        }

        /// <summary>
        /// 将勾选状态写回 SelectedValues（全选时清空集合表示"全部"）
        /// </summary>
        private void SyncToSelectedValues()
        {
            IList<string> selected = SelectedValues;
            if (selected == null)
            {
                return;
            }
            List<string> checkedNames = _allItems.Where(i => i.IsChecked).Select(i => i.Name).ToList();
            selected.Clear();
            if (checkedNames.Count < _allItems.Count)
            {
                foreach (string name in checkedNames)
                {
                    selected.Add(name);
                }
            }
            SyncSelectAllState();
            UpdateSummary();
            SelectionChanged?.Invoke(this, EventArgs.Empty);
        }

        private void SyncSelectAllState()
        {
            if (PART_SelectAll == null)
            {
                return;
            }
            _updating = true;
            PART_SelectAll.IsChecked = _allItems.Count > 0 && _allItems.All(i => i.IsChecked);
            _updating = false;
        }

        private void UpdateSummary()
        {
            List<string> checkedNames = _allItems.Where(i => i.IsChecked).Select(i => i.Name).ToList();
            if (checkedNames.Count == 0)
            {
                SummaryText = "(无选择)";
            }
            else if (checkedNames.Count == _allItems.Count)
            {
                SummaryText = "(全部)";
            }
            else if (checkedNames.Count == 1)
            {
                SummaryText = checkedNames[0];
            }
            else
            {
                SummaryText = $"({checkedNames.Count}项已选)";
            }
        }
    }
}

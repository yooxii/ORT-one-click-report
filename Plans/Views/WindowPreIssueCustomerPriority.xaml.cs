using CommunityToolkit.Mvvm.ComponentModel;
using ORT一键报告.Plans.ViewModels;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// 同系列客户优先级（前 6 位相同的机种只开一张单，按这里的优先级挑客户代码）：
    /// 优先级用数字表示，**数字越小越优先，可以填相同的数字表示并列**
    /// （并列时再按客户代码、机种名排序），没列出的客户代码排最后。
    /// 客户名称从主数据库的客户代码表带出。保存进数据库，多台电脑共用。
    /// </summary>
    public partial class WindowPreIssueCustomerPriority : Window
    {
        private readonly PreIssueViewModel _vm;
        private readonly Dictionary<string, string> _names;
        private readonly ObservableCollection<PriorityRow> _rows = [];

        public WindowPreIssueCustomerPriority(PreIssueViewModel vm)
        {
            InitializeComponent();
            _vm = vm;
            _names = _vm.LoadCustomerCodeNames();
            dg_priority.ItemsSource = _rows;
            cb_code.ItemsSource = _vm.LoadCustomerCodeCandidates();
            foreach (CustomerPriorityEntry entry in _vm.LoadCustomerPriority())
            {
                _rows.Add(BuildRow(entry.Code, entry.Priority));
            }
            SortRows();
            ShowSummary();
        }

        /* ###############################  行构造 / 排序  ################################ */

        private PriorityRow BuildRow(string code, int priority)
        {
            string text = (code ?? "").Trim().ToUpperInvariant();
            return new PriorityRow
            {
                Code = text,
                Name = _names.TryGetValue(text, out string name) ? name : "",
                Priority = priority
            };
        }

        /// <summary>按优先级数字从优到劣排序（并列按客户代码）</summary>
        private void SortRows()
        {
            List<PriorityRow> sorted = _rows
                .OrderBy(r => r.Priority)
                .ThenBy(r => r.Code, StringComparer.OrdinalIgnoreCase)
                .ToList();
            for (int index = 0; index < sorted.Count; index++)
            {
                int current = _rows.IndexOf(sorted[index]);
                if (current != index)
                {
                    _rows.Move(current, index);
                }
            }
        }

        private List<CustomerPriorityEntry> CurrentEntries() => _rows
            .Select(r => new CustomerPriorityEntry { Code = r.Code, Priority = r.Priority })
            .ToList();

        /// <summary>底部状态：当前优先级一览（数字小的在前）</summary>
        private void ShowSummary()
        {
            txt_status.Text = _rows.Count == 0
                ? LanguageService.Get("PreIssue_CustomerPriorityEmpty")
                : PreIssueSeriesPriority.Describe(CurrentEntries());
        }

        private void Dg_Priority_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dg_priority.SelectedItem is PriorityRow row)
            {
                txt_status.Text = row.Name.Length > 0
                    ? $"优先级 {row.Priority}：{row.Code} = {row.Name}"
                    : $"优先级 {row.Priority}：{row.Code}";
            }
        }

        /// <summary>改完优先级数字后刷新底部一览（不在编辑中重排，避免行跳来跳去）</summary>
        private void Dg_Priority_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction != DataGridEditAction.Commit)
            {
                return;
            }
            Dispatcher.BeginInvoke(new Action(ShowSummary), System.Windows.Threading.DispatcherPriority.Background);
        }

        private void Cb_Code_SelectionChanged(object sender, SelectionChangedEventArgs e) => ShowInputName();

        /// <summary>输入/选中一个代码时，先在状态栏里显示它对应的客户名称</summary>
        private void ShowInputName()
        {
            string code = NormalizedInput();
            if (code.Length == 0)
            {
                return;
            }
            txt_status.Text = _names.TryGetValue(code, out string name)
                ? $"{code} = {name}"
                : string.Format(LanguageService.Get("PreIssue_CustomerPriorityNoName"), code);
        }

        private string NormalizedInput()
        {
            string text = (cb_code.Text ?? "").Trim().ToUpperInvariant();
            int bracket = text.IndexOfAny(['（', '(', ' ']);
            return bracket > 0 ? text.Substring(0, bracket).Trim() : text;
        }

        /* ###############################  增删 / 恢复默认  ################################ */

        private void Btn_Add_Click(object sender, RoutedEventArgs e) => AddCode();

        private void Cb_Code_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter)
            {
                return;
            }
            AddCode();
            e.Handled = true;
        }

        private void AddCode()
        {
            string text = NormalizedInput();
            if (text.Length == 0)
            {
                txt_status.Text = LanguageService.Get("PreIssue_CustomerPriorityAddEmpty");
                return;
            }
            if (_rows.Any(r => string.Equals(r.Code, text, StringComparison.OrdinalIgnoreCase)))
            {
                txt_status.Text = string.Format(LanguageService.Get("PreIssue_CustomerPriorityExists"), text);
                return;
            }
            // 新加的代码默认排在最后（现有最大数字 + 1），之后可以改成想要的值
            int next = _rows.Count == 0 ? 1 : _rows.Max(r => r.Priority) + 1;
            PriorityRow row = BuildRow(text, next);
            _rows.Add(row);
            SortRows();
            dg_priority.SelectedItem = row;
            dg_priority.ScrollIntoView(row);
            cb_code.Text = "";
            ShowSummary();
        }

        private void Btn_Remove_Click(object sender, RoutedEventArgs e)
        {
            if (dg_priority.SelectedItem is not PriorityRow row)
            {
                txt_status.Text = LanguageService.Get("PreIssue_CustomerPrioritySelectFirst");
                return;
            }
            _rows.Remove(row);
            ShowSummary();
        }

        private void Btn_Reset_Click(object sender, RoutedEventArgs e)
        {
            _rows.Clear();
            foreach (CustomerPriorityEntry entry in PreIssueSeriesPriority.DefaultEntries())
            {
                _rows.Add(BuildRow(entry.Code, entry.Priority));
            }
            SortRows();
            ShowSummary();
        }

        /* ###############################  保存  ################################ */

        private void Btn_Ok_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 优先级只允许 0 / 正整数：负数没有意义，容易把「最优先」写成 -1 让人看不出来
                PriorityRow invalid = _rows.FirstOrDefault(r => r.Priority < 0);
                if (invalid != null)
                {
                    txt_status.Text = string.Format(LanguageService.Get("PreIssue_CustomerPriorityInvalid"), invalid.Code);
                    dg_priority.SelectedItem = invalid;
                    dg_priority.ScrollIntoView(invalid);
                    return;
                }
                _vm.SaveCustomerPriority(CurrentEntries());
                DialogResult = true;
            }
            catch (Exception ex)
            {
                _ = MessageBox.Show(ex.Message, LanguageService.Get("PreIssue_CustomerPriority"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }

    /// <summary>优先级列表里的一行（优先级数字 / 客户代码 / 客户名称）</summary>
    public class PriorityRow : ObservableObject
    {
        private int _priority;
        private string _code;
        private string _name;

        /// <summary>优先级：数字越小越优先，相同即并列</summary>
        public int Priority
        {
            get => _priority;
            set => SetProperty(ref _priority, value);
        }

        /// <summary>客户代码（机种后段前 2 位）</summary>
        public string Code
        {
            get => _code;
            set => SetProperty(ref _code, value);
        }

        /// <summary>客户名称（客户代码表里查，查不到为空）</summary>
        public string Name
        {
            get => _name;
            set => SetProperty(ref _name, value);
        }
    }
}

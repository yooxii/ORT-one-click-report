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
    /// 同系列客户优先级（前 6 位相同的机种只开一张单，按这里的顺序挑客户代码）：
    /// 列表前面的优先，默认 GA → 4C → 93 → 00（00 = 标准，排最后）；没列出的代码也排最后。
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
            foreach (string code in _vm.LoadCustomerPriorityCodes())
            {
                _rows.Add(BuildRow(code));
            }
            Renumber();
            ShowOrder();
        }

        /* ###############################  行构造 / 顺序  ################################ */

        private PriorityRow BuildRow(string code)
        {
            string text = (code ?? "").Trim().ToUpperInvariant();
            return new PriorityRow
            {
                Code = text,
                Name = _names.TryGetValue(text, out string name) ? name : ""
            };
        }

        /// <summary>重排序号（顺序 = 优先级，1 最高）</summary>
        private void Renumber()
        {
            for (int index = 0; index < _rows.Count; index++)
            {
                _rows[index].Order = index + 1;
            }
        }

        /// <summary>底部状态：当前完整顺序</summary>
        private void ShowOrder()
        {
            txt_status.Text = _rows.Count == 0
                ? LanguageService.Get("PreIssue_CustomerPriorityEmpty")
                : string.Join(" → ", _rows.Select(r => r.Code));
        }

        private void Dg_Priority_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dg_priority.SelectedItem is PriorityRow row)
            {
                txt_status.Text = row.Name.Length > 0
                    ? $"第 {row.Order} 优先：{row.Code} = {row.Name}"
                    : $"第 {row.Order} 优先：{row.Code}";
            }
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

        /* ###############################  增删改 / 排序  ################################ */

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
            PriorityRow row = BuildRow(text);
            _rows.Add(row);
            Renumber();
            dg_priority.SelectedItem = row;
            dg_priority.ScrollIntoView(row);
            cb_code.Text = "";
            ShowOrder();
        }

        private void Btn_Remove_Click(object sender, RoutedEventArgs e)
        {
            if (dg_priority.SelectedItem is not PriorityRow row)
            {
                txt_status.Text = LanguageService.Get("PreIssue_CustomerPrioritySelectFirst");
                return;
            }
            _rows.Remove(row);
            Renumber();
            ShowOrder();
        }

        private void Btn_Up_Click(object sender, RoutedEventArgs e) => Move(-1);

        private void Btn_Down_Click(object sender, RoutedEventArgs e) => Move(1);

        private void Move(int delta)
        {
            if (dg_priority.SelectedItem is not PriorityRow row)
            {
                txt_status.Text = LanguageService.Get("PreIssue_CustomerPrioritySelectFirst");
                return;
            }
            int from = _rows.IndexOf(row);
            int to = from + delta;
            if (to < 0 || to >= _rows.Count)
            {
                return;
            }
            _rows.Move(from, to);
            Renumber();
            dg_priority.SelectedItem = row;
            ShowOrder();
        }

        private void Btn_Reset_Click(object sender, RoutedEventArgs e)
        {
            _rows.Clear();
            foreach (string code in PreIssueSeriesPriority.DefaultCodes)
            {
                _rows.Add(BuildRow(code));
            }
            Renumber();
            ShowOrder();
        }

        private void Btn_Ok_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _vm.SaveCustomerPriorityCodes(_rows.Select(r => r.Code));
                DialogResult = true;
            }
            catch (Exception ex)
            {
                _ = MessageBox.Show(ex.Message, LanguageService.Get("PreIssue_CustomerPriority"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }

    /// <summary>优先级列表里的一行（顺序 / 客户代码 / 客户名称）</summary>
    public class PriorityRow : ObservableObject
    {
        private int _order;
        private string _code;
        private string _name;

        /// <summary>顺序（1 = 最优先）</summary>
        public int Order
        {
            get => _order;
            set => SetProperty(ref _order, value);
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

using NLog;
using ORT一键报告.Models;
using ORT一键报告.Services;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// 计划表批量标记删除窗口（非模态）：工具菜单「批量删除」在计划表页签下打开，
    /// 与主窗口并存——计划表里左键单击一条记录，这里就新增一行（同一条只加一次）。
    /// 清单列出将被删除的全部关键字段供核对；右键单击任一单元格即复制该值。
    /// 点「确认」时通过 <see cref="Confirmed"/> 交给调用方（WindowPlans）标记删除，
    /// 仍走暂存 → 点「提交保存」时统一入库。
    /// </summary>
    public partial class WindowPlanBatchDelete : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly ObservableCollection<PlanBatchRow> _rows = [];
        private readonly HashSet<Plan> _picked = [];

        /// <summary>点了「确认」时触发（调用方负责标记删除并结束会话）</summary>
        public event Action<WindowPlanBatchDelete> Confirmed;

        /// <summary>清单里将被标记删除的计划记录</summary>
        public List<Plan> Targets { get; private set; } = [];

        public WindowPlanBatchDelete()
        {
            InitializeComponent();
            dg_items.ItemsSource = _rows;
            btn_confirm.Content = LanguageService.Get("Plans_MarkDelete");
            RightClickCopy.AttachDataGrid(dg_items, BatchCellValue);
            RefreshSummary();
        }

        /* ###############################  清单增删  ################################ */

        /// <summary>把主窗口计划表里点选的记录加入清单（同一条只加一次）</summary>
        public void AddRecords(IEnumerable<Plan> records)
        {
            if (records == null)
            {
                return;
            }
            bool added = false;
            foreach (Plan plan in records.Where(p => p != null).ToList())
            {
                if (!_picked.Add(plan))
                {
                    continue;
                }
                _rows.Add(new PlanBatchRow(plan, LanguageService.Get("Plans_Batch_StatusOkDelete")));
                added = true;
            }
            if (added)
            {
                RefreshSummary();
            }
        }

        /// <summary>把清单里当前选中的记录移出（点错了可以撤销）</summary>
        private void Btn_Remove_Click(object sender, RoutedEventArgs e)
        {
            List<PlanBatchRow> picked = [.. dg_items.SelectedItems.OfType<PlanBatchRow>()];
            if (picked.Count == 0)
            {
                return;
            }
            foreach (PlanBatchRow row in picked)
            {
                _rows.Remove(row);
                _picked.Remove(row.Item);
            }
            RefreshSummary();
        }

        private void Dg_Items_SelectionChanged(object sender, SelectionChangedEventArgs e)
            => btn_remove.IsEnabled = dg_items.SelectedItems.Count > 0;

        /* ###############################  清单与汇总  ################################ */

        /// <summary>刷新标题、清单条数汇总（含机种分布）与空清单提示</summary>
        private void RefreshSummary()
        {
            Targets = [.. _rows.Select(r => r.Item)];
            Title = string.Format(LanguageService.Get("Plans_Batch_PlanDeleteTitle"), _rows.Count);
            txt_hint.Text = LanguageService.Get("Plans_Batch_PlanDeleteHint");
            txt_summary.Text = string.Format(LanguageService.Get("Plans_Batch_SummaryFormat"),
                _rows.Count, _rows.Count, 0, BuildModelSummary());
            txt_empty.Visibility = _rows.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            btn_confirm.IsEnabled = _rows.Count > 0;
            btn_remove.IsEnabled = dg_items.SelectedItems.Count > 0;
        }

        /// <summary>汇总条里的机种分布（按条数降序）</summary>
        private string BuildModelSummary()
        {
            List<string> parts = [.. _rows
                .GroupBy(r => string.IsNullOrWhiteSpace(r.Item.ModelName) ? "-" : r.Item.ModelName.Trim(), StringComparer.Ordinal)
                .OrderByDescending(g => g.Count())
                .ThenBy(g => g.Key, StringComparer.Ordinal)
                .Select(g => string.Format(LanguageService.Get("Plans_Batch_ModelItemFormat"), g.Key, g.Count()))];
            return parts.Count == 0 ? "-" : string.Join(LanguageService.Get("Plans_Batch_ModelSeparator"), parts);
        }

        /* ###############################  窗口位置与置顶  ################################ */

        /// <summary>显示前定位到主窗口右下角（不遮住表格左上角的记录），用户可自行拖动、缩放</summary>
        public void PlaceNearOwner(Window owner)
        {
            if (owner == null)
            {
                return;
            }
            Owner = owner;
            Rect area = SystemParameters.WorkArea;
            double left = owner.Left + owner.Width - Width - 24;
            double top = owner.Top + owner.Height - Height - 24;
            Left = Math.Max(area.Left, Math.Min(left, area.Right - Width));
            Top = Math.Max(area.Top, Math.Min(top, area.Bottom - Height));
        }

        /// <summary>置顶开关：默认不置顶</summary>
        private void Btn_Topmost_Changed(object sender, RoutedEventArgs e) => Topmost = btn_topmost.IsChecked == true;

        /* ###############################  右键单击复制一次值  ################################ */

        /// <summary>取清单里某个单元格的完整值，供右键复制使用</summary>
        private static string BatchCellValue(object item, DataGridColumn column)
        {
            string field = column?.SortMemberPath;
            if (item is not PlanBatchRow row || string.IsNullOrEmpty(field))
            {
                return "";
            }
            return field == "Result" ? row.Result : PlanFieldText.PlanText(row.Item, field) ?? "";
        }

        /* ###############################  确认 / 取消  ################################ */

        private void Btn_Confirm_Click(object sender, RoutedEventArgs e)
        {
            if (_rows.Count == 0)
            {
                _ = MessageBox.Show(LanguageService.Get("Plans_Batch_EmptyConfirm"), LanguageService.Get("Cap_Info"));
                return;
            }
            _logger.Info($"批量标记删除计划：清单 {_rows.Count} 条");
            Confirmed?.Invoke(this);
            Close();
        }

        /// <summary>取消：整个清单作废（非模态窗口不能设 DialogResult，直接关闭）</summary>
        private void Btn_Cancel_Click(object sender, RoutedEventArgs e) => Close();
    }
}

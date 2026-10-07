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
    /// 清单只列辨认记录所需的信息（工作编号/机种名称/测试项目），右键单击任一单元格即复制该值。
    /// 点「确认」时通过 <see cref="Confirmed"/> 交给调用方（WindowPlans）标记删除，
    /// 仍走暂存 → 点「提交保存」时统一入库。
    /// </summary>
    public partial class WindowPlanBatchDelete : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly ObservableCollection<Plan> _rows = [];
        private readonly HashSet<Plan> _picked = [];

        /// <summary>点了「确认」时触发（调用方负责标记删除并结束会话）</summary>
        public event Action<WindowPlanBatchDelete> Confirmed;

        /// <summary>清单里将被标记删除的计划记录</summary>
        public List<Plan> Targets => [.. _rows];

        public WindowPlanBatchDelete()
        {
            InitializeComponent();
            dg_items.ItemsSource = _rows;
            btn_confirm.Content = LanguageService.Get("Plans_MarkDelete");
            RightClickCopy.AttachDataGrid(dg_items, BatchCellValue);
            RefreshState();
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
                if (_picked.Add(plan))
                {
                    _rows.Add(plan);
                    added = true;
                }
            }
            if (added)
            {
                RefreshState();
            }
        }

        /// <summary>把清单里当前选中的记录移出（点错了可以撤销）</summary>
        private void Btn_Remove_Click(object sender, RoutedEventArgs e)
        {
            List<Plan> picked = [.. dg_items.SelectedItems.OfType<Plan>()];
            // 兜底：整行选中没生效时至少撤掉光标所在的那一行
            if (picked.Count == 0 && dg_items.CurrentItem is Plan current)
            {
                picked.Add(current);
            }
            if (picked.Count == 0)
            {
                return;
            }
            foreach (Plan plan in picked)
            {
                _rows.Remove(plan);
                _picked.Remove(plan);
            }
            RefreshState();
        }

        private void Dg_Items_SelectionChanged(object sender, SelectionChangedEventArgs e)
            => btn_remove.IsEnabled = dg_items.SelectedItems.Count > 0;

        /* ###############################  状态  ################################ */

        private void RefreshState()
        {
            txt_hint.Text = LanguageService.Get("Plans_Batch_PlanDeleteHint");
            Title = string.Format(LanguageService.Get("Plans_Batch_PlanDeleteTitle"), _rows.Count);
            btn_confirm.IsEnabled = _rows.Count > 0;
            btn_remove.IsEnabled = dg_items.SelectedItems.Count > 0;
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
            return item is Plan plan && !string.IsNullOrEmpty(field)
                ? PlanFieldText.PlanText(plan, field) ?? ""
                : "";
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

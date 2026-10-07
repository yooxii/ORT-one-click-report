using NLog;
using ORT一键报告.Models;
using ORT一键报告.Plans.ViewModels;
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
    /// 领退表批量登记窗口（非模态）：工具菜单「批量回线 / 批量删除」打开，
    /// 与主窗口并存——主窗口表格里左键单击一条记录，这里就新增一行（同一条只加一次）。
    /// 清单只列与本次操作有关的信息，口径与单条登记窗口一致：
    /// 回线＝机种名称/回线RT工令/回线数/线别；删除＝领用日期/领料单据号/机种名称。
    /// 右键单击任一单元格即复制该值。
    /// 窗口只收集要写入的值，点「确认」时通过 <see cref="Confirmed"/> 交给
    /// 调用方（WindowPlans）写回，仍走暂存 → 点「提交保存」时统一入库。
    /// </summary>
    public partial class WindowRequisitionBatch : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly RequisitionBatchMode _mode;
        private readonly ObservableCollection<Requisition> _rows = [];
        private readonly HashSet<Requisition> _picked = [];

        /// <summary>点了「确认」时触发（调用方负责把值写回记录并关闭会话）</summary>
        public event Action<WindowRequisitionBatch> Confirmed;

        /// <summary>本次操作类型</summary>
        public RequisitionBatchMode Mode => _mode;

        /// <summary>清单里可执行（会被写入）的记录；不满足条件的不在其中</summary>
        public List<Requisition> Targets
            => [.. _rows.Where(r => RequisitionBatchRules.BlockReason(_mode, r) == null)];

        /// <summary>清单里因不满足条件会被跳过的记录数</summary>
        public int SkippedCount => _rows.Count - Targets.Count;

        /// <summary>确认回线时的日期（回线模式）</summary>
        public DateTime ReturnDate { get; private set; }

        public WindowRequisitionBatch(RequisitionBatchMode mode)
        {
            InitializeComponent();
            _mode = mode;
            dg_items.ItemsSource = _rows;

            // 标题/说明/按钮/输入区/清单列按批量类型切换
            Title = string.Format(LanguageService.Get(TitleKey(mode)), 0);
            txt_hint.Text = LanguageService.Get(HintKey(mode));
            btn_confirm.Content = LanguageService.Get(ConfirmKey(mode));
            // 只有批量回线要填值（批量删除不写值，只标记删除）
            panel_inputs.Visibility = mode == RequisitionBatchMode.Delete ? Visibility.Collapsed : Visibility.Visible;
            ApplyColumns();
            dp_returnDate.SelectedDate = DateTime.Today;

            RightClickCopy.AttachDataGrid(dg_items, BatchCellValue);
            RefreshState();
        }

        /* ###############################  清单列（只留有关信息）  ################################ */

        /// <summary>
        /// 按批量类型显示清单列，口径与单条登记窗口的条目一致：
        /// 回线＝机种名称/回线RT工令/回线数/线别；删除＝领用日期/领料单据号/机种名称
        /// </summary>
        private void ApplyColumns()
        {
            bool isReturn = _mode == RequisitionBatchMode.Return;
            Show(col_reqDate, !isReturn);     // 领用日期：删除时用来辨认记录
            Show(col_reqNo, !isReturn);       // 领料单据号：删除
            Show(col_model, true);            // 机种名称：两种操作都有
            Show(col_returnRt, isReturn);     // 回线RT工令：回线
            Show(col_returnQty, isReturn);    // 回线数量：回线
            Show(col_line, isReturn);         // 线别：回线
        }

        private static void Show(DataGridColumn column, bool visible)
            => column.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;

        /* ###############################  清单增删  ################################ */

        /// <summary>把主窗口表格里点选的记录加入清单（同一条只加一次）</summary>
        public void AddRecords(IEnumerable<Requisition> records)
        {
            if (records == null)
            {
                return;
            }
            bool added = false;
            foreach (Requisition req in records.Where(r => r != null).ToList())
            {
                if (_picked.Add(req))
                {
                    _rows.Add(req);
                    added = true;
                }
            }
            if (added)
            {
                PrefillInputs();
                RefreshState();
            }
        }

        /// <summary>把清单里当前选中的记录移出（点错了可以撤销）</summary>
        private void Btn_Remove_Click(object sender, RoutedEventArgs e)
        {
            List<Requisition> picked = [.. dg_items.SelectedItems.OfType<Requisition>()];
            // 兜底：整行选中没生效时至少撤掉光标所在的那一行
            if (picked.Count == 0 && dg_items.CurrentItem is Requisition current)
            {
                picked.Add(current);
            }
            if (picked.Count == 0)
            {
                return;
            }
            foreach (Requisition req in picked)
            {
                _rows.Remove(req);
                _picked.Remove(req);
            }
            PrefillInputs();
            RefreshState();
        }

        private void Dg_Items_SelectionChanged(object sender, SelectionChangedEventArgs e)
            => btn_remove.IsEnabled = dg_items.SelectedItems.Count > 0;

        /* ###############################  预填要写入的值  ################################ */

        /// <summary>
        /// 清单里已有相同的回线日期时沿用（方便修正），否则用今天
        /// </summary>
        private void PrefillInputs()
        {
            if (_mode == RequisitionBatchMode.Return)
            {
                dp_returnDate.SelectedDate = CommonDate(Targets.Select(r => r.ReturnDate)) ?? DateTime.Today;
            }
        }

        /// <summary>所有值都相同（且非空）时返回该日期，否则 null</summary>
        private static DateTime? CommonDate(IEnumerable<DateTime?> values)
        {
            List<DateTime?> list = [.. values];
            DateTime? first = list.FirstOrDefault();
            return first != null && list.All(v => v == first) ? first : null;
        }

        /* ###############################  状态  ################################ */

        /// <summary>输入或清单变化时刷新标题里的条数与确认/移除按钮可用性</summary>
        private void BatchInput_Changed(object sender, RoutedEventArgs e) => RefreshState();

        private void RefreshState()
        {
            if (_rows == null)
            {
                return;
            }
            Title = string.Format(LanguageService.Get(TitleKey(_mode)), _rows.Count);
            btn_confirm.IsEnabled = Targets.Count > 0;
            btn_remove.IsEnabled = dg_items.SelectedItems.Count > 0;
        }

        /* ###############################  窗口位置与置顶  ################################ */

        /// <summary>
        /// 显示前定位到主窗口右下角（不遮住表格左上角的记录），用户可自行拖动、缩放
        /// </summary>
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

        /// <summary>
        /// 取清单里某个单元格的完整值（含未压成单行显示的原始内容），供右键复制使用
        /// </summary>
        private static string BatchCellValue(object item, DataGridColumn column)
        {
            string field = column?.SortMemberPath;
            return item is Requisition req && !string.IsNullOrEmpty(field)
                ? PlanFieldText.RequisitionText(req, field) ?? ""
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
            if (Targets.Count == 0)
            {
                _ = MessageBox.Show(LanguageService.Get("Plans_Msg_BatchNoneApplied"), LanguageService.Get("Cap_Info"));
                return;
            }
            if (_mode == RequisitionBatchMode.Return)
            {
                if (dp_returnDate.SelectedDate is not DateTime date)
                {
                    _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillReturnDate"), LanguageService.Get("Cap_Info"));
                    return;
                }
                ReturnDate = date;
            }
            _logger.Info($"批量{LanguageService.Get(ConfirmKey(_mode))}：清单 {_rows.Count} 条，写入 {Targets.Count} 条，跳过 {SkippedCount} 条");
            Confirmed?.Invoke(this);
            Close();
        }

        /// <summary>取消：整个清单作废（非模态窗口不能设 DialogResult，直接关闭）</summary>
        private void Btn_Cancel_Click(object sender, RoutedEventArgs e) => Close();

        /* ###############################  资源键  ################################ */

        private static string TitleKey(RequisitionBatchMode mode) => mode switch
        {
            RequisitionBatchMode.Delete => "Plans_Batch_DeleteTitle",
            _ => "Plans_Batch_ReturnTitle"
        };

        private static string HintKey(RequisitionBatchMode mode) => mode switch
        {
            RequisitionBatchMode.Delete => "Plans_Batch_DeleteHint",
            _ => "Plans_Batch_ReturnHint"
        };

        private static string ConfirmKey(RequisitionBatchMode mode) => mode switch
        {
            RequisitionBatchMode.Delete => "Plans_MarkDelete",
            _ => "Plans_Menu_ReturnLine"
        };
    }
}

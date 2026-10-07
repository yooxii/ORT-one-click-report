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
    /// 领退表批量登记窗口（非模态）：工具菜单「批量回线 / 批量入库 / 批量删除」打开，
    /// 与主窗口并存——主窗口表格里左键单击一条记录，这里就新增一行（同一条只加一次）。
    /// 清单把记录的关键字段全部列出（领用、回线、入库、报废、备注），
    /// 「本次结果」列逐条给出将写入什么或跳过原因；右键单击任一单元格即复制该值。
    /// 窗口只收集要写入的值并算出可执行清单，点「确认」时通过 <see cref="Confirmed"/> 交给
    /// 调用方（WindowPlans）写回，仍走暂存 → 点「提交保存」时统一入库。
    /// </summary>
    public partial class WindowRequisitionBatch : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly RequisitionBatchMode _mode;
        private readonly ObservableCollection<RequisitionBatchRow> _rows = [];
        private readonly HashSet<Requisition> _picked = [];

        /// <summary>点了「确认」时触发（调用方负责把值写回记录并关闭会话）</summary>
        public event Action<WindowRequisitionBatch> Confirmed;

        /// <summary>本次操作类型</summary>
        public RequisitionBatchMode Mode => _mode;

        /// <summary>清单里可执行（会被写入）的记录；跳过的不在其中</summary>
        public List<Requisition> Targets { get; private set; } = [];

        /// <summary>清单里因不满足条件被跳过的记录数</summary>
        public int SkippedCount { get; private set; }

        /// <summary>确认回线时的日期（回线模式）</summary>
        public DateTime ReturnDate { get; private set; }

        /// <summary>确认入库时的入库单据（入库模式）</summary>
        public string StockInNo { get; private set; }

        /// <summary>确认入库时的入库数量（入库模式）</summary>
        public string StockInQty { get; private set; }

        /// <summary>确认入库时的入库日期（入库模式）</summary>
        public DateTime StockInDate { get; private set; }

        public WindowRequisitionBatch(RequisitionBatchMode mode)
        {
            InitializeComponent();
            _mode = mode;
            dg_items.ItemsSource = _rows;

            // 标题/说明/按钮/输入区按批量类型切换
            Title = string.Format(LanguageService.Get(TitleKey(mode)), 0);
            txt_hint.Text = LanguageService.Get(HintKey(mode));
            btn_confirm.Content = LanguageService.Get(ConfirmKey(mode));
            panel_return.Visibility = mode == RequisitionBatchMode.Return ? Visibility.Visible : Visibility.Collapsed;
            panel_stockin.Visibility = mode == RequisitionBatchMode.StockIn ? Visibility.Visible : Visibility.Collapsed;
            panel_inputs.Visibility = mode == RequisitionBatchMode.Delete ? Visibility.Collapsed : Visibility.Visible;
            dp_returnDate.SelectedDate = DateTime.Today;
            dp_stockInDate.SelectedDate = DateTime.Today;

            RightClickCopy.AttachDataGrid(dg_items, BatchCellValue);
            RefreshSummary();
        }

        /* ###############################  清单增删  ################################ */

        /// <summary>
        /// 把主窗口表格里点选的记录加入清单（同一条只加一次；不满足条件的也会加入并标出跳过原因）
        /// </summary>
        public void AddRecords(IEnumerable<Requisition> records)
        {
            if (records == null)
            {
                return;
            }
            bool added = false;
            foreach (Requisition req in records.Where(r => r != null).ToList())
            {
                if (!_picked.Add(req))
                {
                    continue;
                }
                _rows.Add(new RequisitionBatchRow(_mode, req));
                added = true;
            }
            if (added)
            {
                PrefillInputs();
                RefreshSummary();
            }
        }

        /// <summary>把清单里当前选中的记录移出（点错了可以撤销）</summary>
        private void Btn_Remove_Click(object sender, RoutedEventArgs e)
        {
            List<RequisitionBatchRow> picked = [.. dg_items.SelectedItems.OfType<RequisitionBatchRow>()];
            if (picked.Count == 0)
            {
                return;
            }
            foreach (RequisitionBatchRow row in picked)
            {
                _rows.Remove(row);
                _picked.Remove(row.Item);
            }
            PrefillInputs();
            RefreshSummary();
        }

        private void Dg_Items_SelectionChanged(object sender, SelectionChangedEventArgs e)
            => btn_remove.IsEnabled = dg_items.SelectedItems.Count > 0;

        /* ###############################  预填要写入的值  ################################ */

        /// <summary>
        /// 清单里已有相同的回线日期/入库信息时沿用（方便修正），否则回线用今天、
        /// 入库日期用今天；单据/数量为空则不预填
        /// </summary>
        private void PrefillInputs()
        {
            List<Requisition> eligible = [.. _rows.Where(r => r.Eligible).Select(r => r.Item)];
            if (_mode == RequisitionBatchMode.Return)
            {
                dp_returnDate.SelectedDate = CommonDate(eligible.Select(r => r.ReturnDate)) ?? DateTime.Today;
                return;
            }
            if (_mode == RequisitionBatchMode.StockIn)
            {
                txt_stockInNo.Text = CommonText(eligible.Select(r => r.StockInNo)) ?? "";
                txt_stockInQty.Text = CommonText(eligible.Select(r => r.StockInQty)) ?? "";
                dp_stockInDate.SelectedDate = CommonDate(eligible.Select(r => r.StockInDate)) ?? DateTime.Today;
            }
        }

        /// <summary>所有值都相同（且非空）时返回该日期，否则 null</summary>
        private static DateTime? CommonDate(IEnumerable<DateTime?> values)
        {
            List<DateTime?> list = [.. values];
            DateTime? first = list.FirstOrDefault();
            return first != null && list.All(v => v == first) ? first : null;
        }

        /// <summary>所有值都相同（且非空）时返回该文本，否则 null</summary>
        private static string CommonText(IEnumerable<string> values)
        {
            List<string> list = [.. values.Where(v => !string.IsNullOrWhiteSpace(v))];
            if (list.Count == 0)
            {
                return null;
            }
            string first = list[0].Trim();
            return list.All(v => string.Equals(v.Trim(), first, StringComparison.Ordinal)) ? first : null;
        }

        /* ###############################  清单与汇总  ################################ */

        /// <summary>输入变化时刷新「本次结果」与汇总，并同步确认按钮可用性</summary>
        private void BatchInput_Changed(object sender, RoutedEventArgs e) => RefreshSummary();

        /// <summary>
        /// 刷新标题、清单条数汇总（含机种分布）与「本次结果」列（将写入什么 / 跳过原因）
        /// </summary>
        private void RefreshSummary()
        {
            if (_rows == null)
            {
                return;
            }
            Targets = [.. _rows.Where(r => r.Eligible).Select(r => r.Item)];
            SkippedCount = _rows.Count - Targets.Count;
            string returnText = ReturnDateText();
            string stockInText = StockInText();
            foreach (RequisitionBatchRow row in _rows)
            {
                if (!row.Eligible)
                {
                    row.Status = string.Format(LanguageService.Get("Plans_Batch_StatusSkipFormat"), row.BlockReason);
                    continue;
                }
                row.Status = _mode switch
                {
                    RequisitionBatchMode.Return => string.Format(LanguageService.Get("Plans_Batch_StatusOkReturnFormat"), returnText),
                    RequisitionBatchMode.StockIn => string.Format(LanguageService.Get("Plans_Batch_StatusOkStockInFormat"), stockInText),
                    _ => LanguageService.Get("Plans_Batch_StatusOkDelete")
                };
            }
            Title = string.Format(LanguageService.Get(TitleKey(_mode)), _rows.Count);
            txt_summary.Text = string.Format(LanguageService.Get("Plans_Batch_SummaryFormat"),
                _rows.Count, Targets.Count, SkippedCount, BuildModelSummary());
            txt_empty.Visibility = _rows.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            btn_confirm.IsEnabled = Targets.Count > 0;
            btn_remove.IsEnabled = dg_items.SelectedItems.Count > 0;
        }

        /// <summary>汇总条里的机种分布（按条数降序，如「A ×2、B ×1」）</summary>
        private string BuildModelSummary()
        {
            List<string> parts = [.. _rows
                .GroupBy(r => string.IsNullOrWhiteSpace(r.Item.ModelName) ? "-" : r.Item.ModelName.Trim(), StringComparer.Ordinal)
                .OrderByDescending(g => g.Count())
                .ThenBy(g => g.Key, StringComparer.Ordinal)
                .Select(g => string.Format(LanguageService.Get("Plans_Batch_ModelItemFormat"), g.Key, g.Count()))];
            return parts.Count == 0 ? "-" : string.Join(LanguageService.Get("Plans_Batch_ModelSeparator"), parts);
        }

        /// <summary>当前回线日期文本（未选日期时给出提示文字）</summary>
        private string ReturnDateText()
            => dp_returnDate.SelectedDate is DateTime date ? date.ToString("yyyy/M/d") : LanguageService.Get("Plans_Batch_NotFilled");

        /// <summary>当前入库信息文本「单据 / 数量 / 日期」</summary>
        private string StockInText()
        {
            string no = string.IsNullOrWhiteSpace(txt_stockInNo.Text) ? LanguageService.Get("Plans_Batch_NotFilled") : txt_stockInNo.Text.Trim();
            string qty = string.IsNullOrWhiteSpace(txt_stockInQty.Text) ? LanguageService.Get("Plans_Batch_NotFilled") : txt_stockInQty.Text.Trim();
            string date = dp_stockInDate.SelectedDate is DateTime d ? d.ToString("yyyy/M/d") : LanguageService.Get("Plans_Batch_NotFilled");
            return $"{no} / {qty} / {date}";
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
            if (item is not RequisitionBatchRow row || string.IsNullOrEmpty(field))
            {
                return "";
            }
            return field == "Status" ? row.Status : PlanFieldText.RequisitionText(row.Item, field) ?? "";
        }

        /* ###############################  确认 / 取消  ################################ */

        private void Btn_Confirm_Click(object sender, RoutedEventArgs e)
        {
            if (_rows.Count == 0)
            {
                _ = MessageBox.Show(LanguageService.Get("Plans_Batch_EmptyConfirm"), LanguageService.Get("Cap_Info"));
                return;
            }
            RefreshSummary();
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
            else if (_mode == RequisitionBatchMode.StockIn)
            {
                if (string.IsNullOrWhiteSpace(txt_stockInNo.Text))
                {
                    _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillStockInNo"), LanguageService.Get("Cap_Info"));
                    txt_stockInNo.Focus();
                    return;
                }
                if (string.IsNullOrWhiteSpace(txt_stockInQty.Text))
                {
                    _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillStockInQty"), LanguageService.Get("Cap_Info"));
                    txt_stockInQty.Focus();
                    return;
                }
                if (dp_stockInDate.SelectedDate is not DateTime date)
                {
                    _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillStockInDate"), LanguageService.Get("Cap_Info"));
                    return;
                }
                StockInNo = txt_stockInNo.Text.Trim();
                StockInQty = txt_stockInQty.Text.Trim();
                StockInDate = date;
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
            RequisitionBatchMode.StockIn => "Plans_Batch_StockInTitle",
            RequisitionBatchMode.Delete => "Plans_Batch_DeleteTitle",
            _ => "Plans_Batch_ReturnTitle"
        };

        private static string HintKey(RequisitionBatchMode mode) => mode switch
        {
            RequisitionBatchMode.StockIn => "Plans_Batch_StockInHint",
            RequisitionBatchMode.Delete => "Plans_Batch_DeleteHint",
            _ => "Plans_Batch_ReturnHint"
        };

        private static string ConfirmKey(RequisitionBatchMode mode) => mode switch
        {
            RequisitionBatchMode.StockIn => "Plans_Menu_StockIn",
            RequisitionBatchMode.Delete => "Plans_MarkDelete",
            _ => "Plans_Menu_ReturnLine"
        };
    }
}

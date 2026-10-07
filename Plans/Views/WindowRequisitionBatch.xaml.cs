using NLog;
using ORT一键报告.Models;
using ORT一键报告.Plans.ViewModels;
using ORT一键报告.Services;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// 领退表批量登记窗口：多选记录后右键「回线 / 入库 / 标记删除」共用。
    /// 清单把选中记录的关键字段全部列出（领用、回线、入库、报废、备注），
    /// 「本次结果」列逐条给出将写入什么或跳过原因；右键单击任一单元格即复制该值。
    /// 窗口只收集要写入的值并算出可执行清单，实际写回由调用方（WindowPlans）完成，
    /// 仍走暂存 → 点「提交保存」时统一入库。
    /// </summary>
    public partial class WindowRequisitionBatch : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly RequisitionBatchMode _mode;
        private readonly List<RequisitionBatchRow> _rows;

        /// <summary>本次操作类型</summary>
        public RequisitionBatchMode Mode => _mode;

        /// <summary>可执行（会被写入）的记录；跳过的不在其中</summary>
        public List<Requisition> Targets { get; private set; } = [];

        /// <summary>因不满足条件被跳过的记录数</summary>
        public int SkippedCount { get; private set; }

        /// <summary>确认回线时的日期（回线模式，DialogResult 为 true 时有效）</summary>
        public DateTime ReturnDate { get; private set; }

        /// <summary>确认入库时的入库单据（入库模式，DialogResult 为 true 时有效）</summary>
        public string StockInNo { get; private set; }

        /// <summary>确认入库时的入库数量（入库模式，DialogResult 为 true 时有效）</summary>
        public string StockInQty { get; private set; }

        /// <summary>确认入库时的入库日期（入库模式，DialogResult 为 true 时有效）</summary>
        public DateTime StockInDate { get; private set; }

        public WindowRequisitionBatch(RequisitionBatchMode mode, IEnumerable<Requisition> records)
        {
            InitializeComponent();
            _mode = mode;
            List<Requisition> list = records?.Where(r => r != null).ToList() ?? [];
            _rows = [.. list.Select(r => new RequisitionBatchRow(mode, r))];
            dg_items.ItemsSource = _rows;
            RecalcTargets();

            // 标题/说明/按钮/输入区按批量类型切换
            Title = string.Format(LanguageService.Get(TitleKey(mode)), _rows.Count);
            txt_hint.Text = LanguageService.Get(HintKey(mode));
            btn_confirm.Content = LanguageService.Get(ConfirmKey(mode));
            panel_return.Visibility = mode == RequisitionBatchMode.Return ? Visibility.Visible : Visibility.Collapsed;
            panel_stockin.Visibility = mode == RequisitionBatchMode.StockIn ? Visibility.Visible : Visibility.Collapsed;
            panel_inputs.Visibility = mode == RequisitionBatchMode.Delete ? Visibility.Collapsed : Visibility.Visible;

            PrefillInputs();
            RightClickCopy.AttachDataGrid(dg_items, BatchCellValue);
            RefreshStatus();
        }

        /* ###############################  预填要写入的值  ################################ */

        /// <summary>
        /// 可执行记录已有相同的回线日期/入库信息时沿用（方便修正），否则回线用今天、
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
                string no = CommonText(eligible.Select(r => r.StockInNo));
                string qty = CommonText(eligible.Select(r => r.StockInQty));
                txt_stockInNo.Text = no ?? "";
                txt_stockInQty.Text = qty ?? "";
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

        /// <summary>重新计算可执行清单与跳过条数（按当前选中记录的可执行性）</summary>
        private void RecalcTargets()
        {
            Targets = [.. _rows.Where(r => r.Eligible).Select(r => r.Item)];
            SkippedCount = _rows.Count - Targets.Count;
        }

        /// <summary>输入变化时刷新「本次结果」与汇总，并同步确认按钮可用性</summary>
        private void BatchInput_Changed(object sender, RoutedEventArgs e) => RefreshStatus();

        /// <summary>
        /// 刷新「本次结果」列（将写入什么 / 跳过原因）与顶部汇总（条数、机种分布）
        /// </summary>
        private void RefreshStatus()
        {
            if (_rows == null)
            {
                return;
            }
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
            txt_summary.Text = string.Format(LanguageService.Get("Plans_Batch_SummaryFormat"),
                _rows.Count, Targets.Count, SkippedCount, BuildModelSummary());
            btn_confirm.IsEnabled = Targets.Count > 0;
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
            RecalcTargets();
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
            _logger.Info($"批量{LanguageService.Get(ConfirmKey(_mode))}：写入 {Targets.Count} 条，跳过 {SkippedCount} 条");
            DialogResult = true;
        }

        private void Btn_Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;

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

using NLog;
using ORT一键报告.Models;
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
    /// 计划表批量标记删除窗口：多选计划记录后右键「标记删除」使用。
    /// 清单列出将被删除的全部关键字段供核对，「本次结果」列统一为将标记删除；
    /// 右键单击任一单元格即复制该值。窗口只做确认，实际标记删除由调用方（WindowPlans）完成，
    /// 仍走暂存 → 点「提交保存」时统一入库。
    /// </summary>
    public partial class WindowPlanBatchDelete : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        /// <summary>确认后将被标记删除的计划记录</summary>
        public List<Plan> Targets { get; } = [];

        public WindowPlanBatchDelete(IEnumerable<Plan> records)
        {
            InitializeComponent();
            List<Plan> list = [.. (records ?? []).Where(p => p != null)];
            Targets = list;
            string result = LanguageService.Get("Plans_Batch_StatusOkDelete");
            List<PlanBatchRow> rows = [.. list.Select(p => new PlanBatchRow(p, result))];
            dg_items.ItemsSource = rows;
            Title = string.Format(LanguageService.Get("Plans_Batch_PlanDeleteTitle"), list.Count);
            txt_hint.Text = LanguageService.Get("Plans_Batch_PlanDeleteHint");
            btn_confirm.Content = LanguageService.Get("Plans_MarkDelete");
            txt_summary.Text = string.Format(LanguageService.Get("Plans_Batch_SummaryFormat"),
                list.Count, list.Count, 0, BuildModelSummary(list));
            btn_confirm.IsEnabled = list.Count > 0;
            RightClickCopy.AttachDataGrid(dg_items, BatchCellValue);
        }

        /// <summary>汇总条里的机种分布（按条数降序）</summary>
        private static string BuildModelSummary(IEnumerable<Plan> plans)
        {
            List<string> parts = [.. plans
                .GroupBy(p => string.IsNullOrWhiteSpace(p.ModelName) ? "-" : p.ModelName.Trim(), StringComparer.Ordinal)
                .OrderByDescending(g => g.Count())
                .ThenBy(g => g.Key, StringComparer.Ordinal)
                .Select(g => string.Format(LanguageService.Get("Plans_Batch_ModelItemFormat"), g.Key, g.Count()))];
            return parts.Count == 0 ? "-" : string.Join(LanguageService.Get("Plans_Batch_ModelSeparator"), parts);
        }

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

        private void Btn_Confirm_Click(object sender, RoutedEventArgs e)
        {
            if (Targets.Count == 0)
            {
                _ = MessageBox.Show(LanguageService.Get("Plans_Msg_BatchNoneApplied"), LanguageService.Get("Cap_Info"));
                return;
            }
            _logger.Info($"批量标记删除计划：{Targets.Count} 条");
            DialogResult = true;
        }

        private void Btn_Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;
    }
}

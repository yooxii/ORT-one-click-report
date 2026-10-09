using ORT一键报告.Plans.ViewModels;
using ORT一键报告.Services;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// 回线工令编号检查窗口（只读）：工具菜单打开，列出领退表里「回线RT工令」编号有问题的记录——
    /// 月段与领用日期所在年月不一致（如 1 月的记录写成 RTAH2610…，会让该年月的新编号从错误序号往后接），
    /// 以及根本不是编号形状的值。只做提示，不修改任何数据；右键单击任一单元格即复制该值。
    /// </summary>
    public partial class WindowReturnRtCheck : Window
    {
        /// <summary>清单行：把问题类型翻成界面文案（本地化留在界面层，规则层只给枚举）</summary>
        public sealed class IssueRow
        {
            public DateTime? RequisitionDate { get; set; }
            public DateTime? ReturnDate { get; set; }
            public string RequisitionNo { get; set; }
            public string ModelName { get; set; }
            public string ReturnRtOrder { get; set; }
            public string SuggestedCode { get; set; }
            public string IssueText { get; set; }
        }

        private readonly List<IssueRow> _rows;

        public WindowReturnRtCheck(IReadOnlyList<ReturnRtCodeIssue> issues, int scannedCount)
        {            InitializeComponent();
            _rows = [.. issues.Select(ToRow)];
            dg_issues.ItemsSource = _rows;

            txt_summary.Text = string.Format(
                LanguageService.Get("ReturnRtCheck_Summary"), scannedCount, _rows.Count);
            // 没有问题时不显示空表格，只留一句「没有发现问题」
            bool clean = _rows.Count == 0;
            dg_issues.Visibility = clean ? Visibility.Collapsed : Visibility.Visible;
            txt_clean.Visibility = clean ? Visibility.Visible : Visibility.Collapsed;
            btn_copyAll.IsEnabled = !clean;

            RightClickCopy.AttachDataGrid(dg_issues, CellValue);
        }

        /// <summary>把规则层的结果翻成界面行</summary>
        private static IssueRow ToRow(ReturnRtCodeIssue issue) => new()
        {
            RequisitionDate = issue.RequisitionDate,
            ReturnDate = issue.ReturnDate,
            RequisitionNo = issue.RequisitionNo,
            ModelName = issue.ModelName,
            ReturnRtOrder = issue.ReturnRtOrder,
            SuggestedCode = issue.SuggestedCode,
            IssueText = IssueText(issue)
        };

        /// <summary>问题描述：带上编号里实际写的年月，方便直接核对</summary>
        private static string IssueText(ReturnRtCodeIssue issue) => issue.Kind switch
        {
            ReturnRtCodeIssueKind.MonthMismatch => string.Format(
                LanguageService.Get("ReturnRtCheck_MonthMismatch"), issue.CodeYearMonth),
            _ => LanguageService.Get("ReturnRtCheck_BadFormat")
        };

        /// <summary>取某个单元格的完整值，供右键复制使用</summary>
        private static string CellValue(object item, DataGridColumn column)
        {
            if (item is not IssueRow row)
            {
                return "";
            }
            return column?.SortMemberPath switch
            {
                "RequisitionDate" => row.RequisitionDate?.ToString("yyyy/M/d") ?? "",
                "ReturnDate" => row.ReturnDate?.ToString("yyyy/M/d") ?? "",
                "RequisitionNo" => row.RequisitionNo ?? "",
                "ModelName" => row.ModelName ?? "",
                "ReturnRtOrder" => row.ReturnRtOrder ?? "",
                "SuggestedCode" => row.SuggestedCode ?? "",
                "IssueText" => row.IssueText ?? "",
                _ => ""
            };
        }

        /* ###############################  事件函数  ################################ */

        /// <summary>把整份清单按制表符分隔复制到剪贴板，可直接粘进 Excel 与源表核对</summary>
        private void Btn_CopyAll_Click(object sender, RoutedEventArgs e)
        {
            if (_rows.Count == 0)
            {
                return;
            }
            StringBuilder text = new();
            text.AppendLine(string.Join("\t", LanguageService.Get("Plans_RequisitionDateTC"),
                LanguageService.Get("Plans_RequisitionDocNo"), LanguageService.Get("Plans_ModelNameTC"),
                LanguageService.Get("Plans_ReturnRTOrder"), LanguageService.Get("ReturnRtCheck_Suggested"),
                LanguageService.Get("ReturnRtCheck_Issue")));
            foreach (IssueRow row in _rows)
            {
                text.AppendLine(string.Join("\t", row.RequisitionDate?.ToString("yyyy/M/d") ?? "",
                    row.RequisitionNo, row.ModelName, row.ReturnRtOrder, row.SuggestedCode, row.IssueText));
            }
            try
            {
                Clipboard.SetText(text.ToString());
                ToastService.Show(string.Format(LanguageService.Get("ReturnRtCheck_Copied"), _rows.Count));
            }
            catch (Exception ex)
            {
                _ = MessageBox.Show(ex.Message, LanguageService.Get("Cap_Error"));
            }
        }

        private void Btn_Close_Click(object sender, RoutedEventArgs e) => Close();
    }
}
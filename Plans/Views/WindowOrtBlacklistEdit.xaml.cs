using ORT一键报告.Models;
using ORT一键报告.Plans.ViewModels;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// 黑名单条目编辑框：加入黑名单前先让用户看一眼、改一改，点「完成」才真正写入。
    /// 还能把单个机种一键转成「同系列」正则（前 6 位相同即同一系列，后 4 位 = 客户短码 2 位 + 版本 2 位）。
    /// </summary>
    public partial class WindowOrtBlacklistEdit : Window
    {
        private readonly PreIssueViewModel _vm;
        private bool _initialized;

        /// <summary>确认后要加入的黑名单内容</summary>
        public string Value { get; private set; }

        /// <summary>是否正则模式</summary>
        public bool IsPattern { get; private set; }

        /// <summary>加入原因（可为空）</summary>
        public string Reason { get; private set; }

        public WindowOrtBlacklistEdit(PreIssueViewModel vm, string model, bool isPattern = false, string reason = null)
        {
            InitializeComponent();
            _vm = vm;
            List<KindOption> kinds =
            [
                new KindOption { IsPattern = false, Text = PreIssueCodes.BlacklistKindExact },
                new KindOption { IsPattern = true, Text = PreIssueCodes.BlacklistKindPattern }
            ];
            cb_kind.ItemsSource = kinds;
            cb_kind.SelectedIndex = isPattern ? 1 : 0;
            txt_value.Text = model ?? "";
            txt_reason.Text = reason ?? "";
            _initialized = true;
            UpdateSeriesButton();
            UpdatePreview();
            Loaded += (s, e) =>
            {
                txt_value.Focus();
                txt_value.SelectAll();
            };
        }

        private bool IsPatternSelected => cb_kind.SelectedItem is KindOption option && option.IsPattern;

        private void Input_Changed(object sender, EventArgs e)
        {
            if (!_initialized)
            {
                return;
            }
            UpdateSeriesButton();
            UpdatePreview();
        }

        /// <summary>「按系列生成正则」按钮：只在当前内容能拆成「前 6 位-后 4 位」时可用</summary>
        private void UpdateSeriesButton()
        {
            if (btn_series == null || txt_series == null)
            {
                return;
            }
            bool canSeries = PreIssueViewModel.TryBuildSeriesPattern(txt_value.Text, out _, out string series,
                out string customer, out string version);
            btn_series.IsEnabled = canSeries;
            txt_series.Text = !canSeries
                ? ""
                : string.IsNullOrEmpty(customer)
                    ? string.Format(LanguageService.Get("PreIssue_BlacklistSeriesOnly"), series)
                    : string.Format(LanguageService.Get("PreIssue_BlacklistSeriesInfo"), series, customer, version);
        }

        private void Btn_Series_Click(object sender, RoutedEventArgs e)
        {
            if (!PreIssueViewModel.TryBuildSeriesPattern(txt_value.Text, out string pattern, out string series,
                    out string customer, out string version))
            {
                _ = MessageBox.Show(LanguageService.Get("PreIssue_BlacklistSeriesUnsupported"), Title,
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            cb_kind.SelectedIndex = 1;   // 切到正则
            txt_value.Text = pattern;
            txt_series.Text = string.IsNullOrEmpty(customer)
                ? string.Format(LanguageService.Get("PreIssue_BlacklistSeriesOnly"), series)
                : string.Format(LanguageService.Get("PreIssue_BlacklistSeriesInfo"), series, customer, version);
            UpdatePreview();
        }

        /// <summary>预览这条规则本月会命中哪些机种（写错的正则这里会给出错误说明）</summary>
        private void UpdatePreview()
        {
            if (txt_preview == null)
            {
                return;
            }
            string value = (txt_value.Text ?? "").Trim();
            bool pattern = IsPatternSelected;
            if (value.Length == 0)
            {
                txt_preview.Text = LanguageService.Get("PreIssue_BlacklistEditHint");
                return;
            }
            if (pattern && !PreIssueViewModel.ValidateBlacklistPattern(value, out string patternError))
            {
                txt_preview.Text = string.Format(LanguageService.Get("PreIssue_BlacklistPatternInvalid"), patternError);
                return;
            }
            List<string> hits = _vm?.PreviewBlacklistRule(value, pattern) ?? [];
            txt_preview.Text = hits.Count == 0
                ? LanguageService.Get("PreIssue_BlacklistPreviewNone")
                : string.Format(LanguageService.Get("PreIssue_BlacklistPreview"),
                    hits.Count, string.Join("、", hits.Take(12)) + (hits.Count > 12 ? "…" : ""));
        }

        private void Btn_Ok_Click(object sender, RoutedEventArgs e)
        {
            string value = (txt_value.Text ?? "").Trim();
            bool pattern = IsPatternSelected;
            if (value.Length == 0)
            {
                _ = MessageBox.Show(pattern
                        ? LanguageService.Get("PreIssue_BlacklistPatternRequired")
                        : LanguageService.Get("PreIssue_ModelRequired"),
                    Title, MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (pattern && !PreIssueViewModel.ValidateBlacklistPattern(value, out string patternError))
            {
                _ = MessageBox.Show(string.Format(LanguageService.Get("PreIssue_BlacklistPatternInvalid"), patternError),
                    Title, MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            Value = value;
            IsPattern = pattern;
            Reason = string.IsNullOrWhiteSpace(txt_reason.Text) ? null : txt_reason.Text.Trim();
            DialogResult = true;
        }

        /// <summary>类型下拉项（機種 / 正则）</summary>
        private sealed class KindOption
        {
            public bool IsPattern { get; set; }
            public string Text { get; set; }
            public override string ToString() => Text;
        }
    }
}

using ORT一键报告.Models;
using ORT一键报告.Plans.ViewModels;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// ORT 黑名单管理：有些机种 ORT 做不了、不需要领用，加进黑名单后看板变动不再推进到准备领用。
    /// 条目支持两种类型——「機種」精确匹配、「正则」按正则包含匹配（可一条挡住一整组机种）。
    /// </summary>
    public partial class WindowOrtBlacklist : Window
    {
        private readonly PreIssueViewModel _vm;

        public WindowOrtBlacklist(PreIssueViewModel vm)
        {
            InitializeComponent();
            _vm = vm;
            List<KindOption> kinds =
            [
                new KindOption { IsPattern = false, Text = PreIssueCodes.BlacklistKindExact },
                new KindOption { IsPattern = true, Text = PreIssueCodes.BlacklistKindPattern }
            ];
            cb_kind.ItemsSource = kinds;
            cb_kind.SelectedIndex = 0;
            Reload();
        }

        private bool IsPatternSelected => cb_kind.SelectedItem is KindOption option && option.IsPattern;

        private void Cb_Kind_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (txt_pattern_hint == null)
            {
                return;   // XAML 还在加载
            }
            bool pattern = IsPatternSelected;
            txt_pattern_hint.Visibility = pattern ? Visibility.Visible : Visibility.Collapsed;
            txt_model.ToolTip = pattern ? txt_pattern_hint.Text : null;
        }

        private void Reload()
        {
            List<OrtBlacklistItem> items = _vm.LoadBlacklist();
            dg_blacklist.ItemsSource = items;
            int patterns = 0;
            foreach (OrtBlacklistItem item in items)
            {
                if (item.IsPattern)
                {
                    patterns++;
                }
            }
            txt_status.Text = items.Count == 0
                ? LanguageService.Get("PreIssue_BlacklistEmpty")
                : string.Format(LanguageService.Get("PreIssue_BlacklistCount"), items.Count, patterns);
        }

        private void Btn_Add_Click(object sender, RoutedEventArgs e)
        {
            string value = (txt_model.Text ?? "").Trim();
            bool pattern = IsPatternSelected;
            if (value.Length == 0)
            {
                _ = MessageBox.Show(pattern ? LanguageService.Get("PreIssue_BlacklistPatternRequired")
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
            if (_vm.AddBlacklistModel(value, (txt_reason.Text ?? "").Trim(), pattern))
            {
                txt_model.Clear();
                txt_reason.Clear();
                Reload();
                txt_status.Text = string.Format(LanguageService.Get("PreIssue_BlacklistAdded"), value);
                _vm.ReloadAfterBlacklistChange();
            }
        }

        private void Btn_Remove_Click(object sender, RoutedEventArgs e)
        {
            if (dg_blacklist.SelectedItem is not OrtBlacklistItem item || string.IsNullOrWhiteSpace(item.Model))
            {
                txt_status.Text = LanguageService.Get("PreIssue_BlacklistSelectFirst");
                return;
            }
            if (_vm.RemoveBlacklistModel(item.Model))
            {
                Reload();
                txt_status.Text = string.Format(LanguageService.Get("PreIssue_BlacklistRemoved"), item.Model);
                _vm.ReloadAfterBlacklistChange();
            }
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

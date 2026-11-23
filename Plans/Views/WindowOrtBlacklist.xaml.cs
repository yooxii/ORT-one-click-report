using ORT一键报告.Models;
using ORT一键报告.Plans.ViewModels;
using ORT一键报告.Services;
using System.Collections.Generic;
using System.Windows;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// ORT 黑名单管理：列出已加入的条目（機種 / 正则），新增走「黑名单条目编辑框」，
    /// 也可以把选中的条目移出黑名单。ORT 做不了、不需要领用的机种加进来后，
    /// 看板变动不再推进到准备领用，刷新时也会把已有的待开单清掉。
    /// </summary>
    public partial class WindowOrtBlacklist : Window
    {
        private readonly PreIssueViewModel _vm;

        public WindowOrtBlacklist(PreIssueViewModel vm)
        {
            InitializeComponent();
            _vm = vm;
            Reload();
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

        /// <summary>新增条目：先打开编辑框让用户再确认/修改，点「完成」才写入</summary>
        private void Btn_New_Click(object sender, RoutedEventArgs e)
        {
            WindowOrtBlacklistEdit dialog = new(_vm, null)
            {
                Owner = this
            };
            if (dialog.ShowDialog() == true)
            {
                if (_vm.AddBlacklistModel(dialog.Value, dialog.Reason, dialog.IsPattern))
                {
                    Reload();
                    txt_status.Text = string.Format(LanguageService.Get("PreIssue_BlacklistAdded"), dialog.Value);
                    _vm.ReloadAfterBlacklistChange();
                }
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
    }
}

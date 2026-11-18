using ORT一键报告.Models;
using ORT一键报告.Plans.ViewModels;
using ORT一键报告.Services;
using System;
using System.Windows;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// ORT 黑名单管理：有些机种 ORT 做不了、不需要领用，加进黑名单后看板变动不再推进到准备领用。
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
            dg_blacklist.ItemsSource = _vm.LoadBlacklist();
        }

        private void Btn_Add_Click(object sender, RoutedEventArgs e)
        {
            string model = (txt_model.Text ?? "").Trim();
            if (model.Length == 0)
            {
                _ = MessageBox.Show(LanguageService.Get("PreIssue_ModelRequired"), Title,
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (_vm.AddBlacklistModel(model, (txt_reason.Text ?? "").Trim()))
            {
                txt_model.Clear();
                txt_reason.Clear();
                Reload();
                txt_status.Text = string.Format(LanguageService.Get("PreIssue_BlacklistAdded"), model);
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
    }
}

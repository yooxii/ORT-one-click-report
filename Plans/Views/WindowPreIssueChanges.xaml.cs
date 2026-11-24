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
    /// 预领的变动记录：三个表每次导入更新都逐条记录（新增 / 删除 / 修改 / 入库 / 开单 / 黑名单），
    /// 旧月份被更新时同样记录，可以随时翻看。
    /// </summary>
    public partial class WindowPreIssueChanges : Window
    {
        private readonly PreIssueViewModel _vm;
        private bool _initialized;

        public WindowPreIssueChanges(PreIssueViewModel vm)
        {
            InitializeComponent();
            _vm = vm;
            InitializeFilters();
            _initialized = true;
            Reload();
        }

        private void InitializeFilters()
        {
            List<MonthOption> months = [new MonthOption { Month = null, Text = LanguageService.Get("Common_All") }];
            months.AddRange(_vm.Months);
            cb_month.ItemsSource = months;
            cb_month.SelectedItem = months.FirstOrDefault(m => m.Month == _vm.SelectedMonth?.Month) ?? months[0];

            cb_source.ItemsSource = new List<string>
            {
                LanguageService.Get("Common_All"),
                PreIssueCodes.SourceSchedule,
                PreIssueCodes.SourceBoard,
                PreIssueCodes.SourceRequisition,
                PreIssueCodes.SourceBorrow,
                PreIssueCodes.SourceBlacklist
            };
            cb_source.SelectedIndex = 0;
        }

        private void Filter_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (_initialized)
            {
                Reload();
            }
        }

        private void Btn_Refresh_Click(object sender, RoutedEventArgs e) => Reload();

        private void Reload()
        {
            try
            {
                List<PreIssueChangeLog> logs;
                string source = cb_source.SelectedIndex <= 0 ? null : cb_source.SelectedItem as string;
                if (cb_month.SelectedItem is MonthOption option && !string.IsNullOrEmpty(option.Month))
                {
                    logs = _vm.LoadChangesForMonth(option.Month, source);
                }
                else
                {
                    logs = _vm.LoadAllChangesForView();
                    if (!string.IsNullOrEmpty(source))
                    {
                        logs = logs.Where(l => l.Source == source).ToList();
                    }
                }
                dg_changes.ItemsSource = logs;
                txt_count.Text = string.Format(LanguageService.Get("PreIssue_ChangesCount"), logs.Count);
                txt_status.Text = logs.Count == 0 ? LanguageService.Get("PreIssue_ChangesEmpty") : "";
            }
            catch (Exception ex)
            {
                txt_status.Text = ex.Message;
            }
        }
    }
}

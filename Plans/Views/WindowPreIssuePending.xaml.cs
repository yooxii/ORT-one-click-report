using ORT一键报告.Services;
using System.Collections.Generic;
using System.Windows;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// 添加待开单：准备领用一开始只知道机种名，这里填机种（可直接从看板出现过的机种里挑）与可选工令。
    /// </summary>
    public partial class WindowPreIssuePending : Window
    {
        /// <summary>要添加的机种</summary>
        public string Model { get; private set; }

        /// <summary>可选：关联工令</summary>
        public string JobNo { get; private set; }

        public WindowPreIssuePending(IEnumerable<string> candidates)
        {
            InitializeComponent();
            List<string> list = [];
            if (candidates != null)
            {
                list.AddRange(candidates);
            }
            cb_model.ItemsSource = list;
            txt_hint.Text = LanguageService.Get("PreIssue_AddPendingHint");
            Loaded += (s, e) => cb_model.Focus();
        }

        private void Btn_Ok_Click(object sender, RoutedEventArgs e)
        {
            string model = (cb_model.Text ?? "").Trim();
            if (model.Length == 0)
            {
                _ = MessageBox.Show(LanguageService.Get("PreIssue_ModelRequired"), Title,
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            Model = model;
            string job = (txt_job.Text ?? "").Trim();
            JobNo = job.Length == 0 ? null : job;
            DialogResult = true;
        }
    }
}

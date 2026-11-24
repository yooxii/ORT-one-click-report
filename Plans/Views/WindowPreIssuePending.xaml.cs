using ORT一键报告.Services;
using System.Collections.Generic;
using System.Windows;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// 添加待开单：准备领用一开始只知道机种名，这里填机种（可直接从看板出现过的机种里挑）、
    /// 可选工令与数量；从月排程右键进来时会预填机种与本月合计数量。
    /// </summary>
    public partial class WindowPreIssuePending : Window
    {
        /// <summary>要添加的机种</summary>
        public string Model { get; private set; }

        /// <summary>可选：关联工令</summary>
        public string JobNo { get; private set; }

        /// <summary>可选：领用数量（填了就写进 REQ_QTY）</summary>
        public string Qty { get; private set; }

        /// <summary>可选：备注（同机种重复领用时写原因，会记进变动记录）</summary>
        public string Note { get; private set; }

        public WindowPreIssuePending(IEnumerable<string> candidates, string model = null, string jobNo = null, string qty = null)
        {
            InitializeComponent();
            List<string> list = [];
            if (candidates != null)
            {
                list.AddRange(candidates);
            }
            cb_model.ItemsSource = list;
            txt_hint.Text = LanguageService.Get("PreIssue_AddPendingHint");
            if (!string.IsNullOrWhiteSpace(model))
            {
                cb_model.Text = model.Trim();
            }
            if (!string.IsNullOrWhiteSpace(jobNo))
            {
                txt_job.Text = jobNo.Trim();
            }
            if (!string.IsNullOrWhiteSpace(qty))
            {
                txt_qty.Text = qty.Trim();
            }
            // 预填了机种的（从月排程右键进来）直接跳到数量，省得再点
            Loaded += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(cb_model.Text))
                {
                    cb_model.Focus();
                }
                else
                {
                    txt_qty.Focus();
                    txt_qty.SelectAll();
                }
            };
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
            string qty = (txt_qty.Text ?? "").Trim();
            Qty = qty.Length == 0 ? null : qty;
            string note = (txt_note.Text ?? "").Trim();
            Note = note.Length == 0 ? null : note;
            DialogResult = true;
        }
    }
}

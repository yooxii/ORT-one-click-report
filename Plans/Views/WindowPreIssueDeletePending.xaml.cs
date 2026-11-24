using ORT一键报告.Services;
using System.Windows;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// 删除待开单的确认框：三个按钮——是（只删除）、是并加入黑名单（删除 + 加入黑名单）、否（不删）。
    /// 常规 MessageBox 不能改按钮文字，所以单独做一个。
    /// </summary>
    public partial class WindowPreIssueDeletePending : Window
    {
        /// <summary>点「是并加入黑名单」时为 true（同时删除并加入黑名单）</summary>
        public bool AddBlacklist { get; private set; }

        public WindowPreIssueDeletePending(string model)
        {
            InitializeComponent();
            txt_ask.Text = string.Format(LanguageService.Get("PreIssue_DeletePendingAsk"), model ?? "");
        }

        private void Btn_Yes_Click(object sender, RoutedEventArgs e)
        {
            AddBlacklist = false;
            DialogResult = true;
        }

        private void Btn_YesBlacklist_Click(object sender, RoutedEventArgs e)
        {
            AddBlacklist = true;
            DialogResult = true;
        }
    }
}

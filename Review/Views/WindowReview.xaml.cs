using Microsoft.Extensions.DependencyInjection;
using ORT一键报告.Models;
using ORT一键报告.Services;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ORT一键报告.Review.Views
{
    /// <summary>
    /// WindowReview.xaml 的交互逻辑：审核工作流主界面（请求列表展开，点击打开详情审核）
    /// </summary>
    public partial class WindowReview : Window
    {
        private readonly ReviewService _reviewService;

        public WindowReview()
        {
            InitializeComponent();
            _reviewService = App.ServiceProvider.GetRequiredService<ReviewService>();
            Loaded += (s, e) => LoadRequests();
        }

        private void LoadRequests()
        {
            string status = (cb_status.SelectedItem as ComboBoxItem)?.Content?.ToString();
            List<ReviewRequest> requests = _reviewService.GetRequests(status == "全部" ? null : status);
            dg_requests.ItemsSource = requests;
            status_msg.Content = $"共 {requests.Count} 条请求";
        }

        private void OpenDetail()
        {
            if (dg_requests.SelectedItem is not ReviewRequest request)
            {
                _ = MessageBox.Show("请先选择一个请求", "提示");
                return;
            }
            WindowReviewDetail detail = new(request)
            {
                Owner = this
            };
            detail.ShowDialog();
            LoadRequests();
        }

        /* ###############################  事件函数  ################################ */

        private void Cb_Status_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dg_requests != null)
            {
                LoadRequests();
            }
        }

        private void Btn_Refresh_Click(object sender, RoutedEventArgs e) => LoadRequests();

        private void Btn_Detail_Click(object sender, RoutedEventArgs e) => OpenDetail();

        private void Dg_Requests_MouseDoubleClick(object sender, MouseButtonEventArgs e) => OpenDetail();
    }
}

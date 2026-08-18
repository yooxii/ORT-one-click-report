using Microsoft.Extensions.DependencyInjection;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using ORT一键报告.Models;
using ORT一键报告.Services;
using System;
using System.Windows;

namespace ORT一键报告.Review.Views
{
    /// <summary>
    /// WindowReviewDetail.xaml 的交互逻辑：请求详情查看与审核（通过/驳回）
    /// </summary>
    public partial class WindowReviewDetail : Window
    {
        private readonly ReviewRequest _request;
        private readonly ReviewService _reviewService;
        private readonly AuthService _auth;

        public WindowReviewDetail(ReviewRequest request)
        {
            InitializeComponent();
            _request = request;
            _reviewService = App.ServiceProvider.GetRequiredService<ReviewService>();
            _auth = App.ServiceProvider.GetRequiredService<AuthService>();

            txt_summary.Text = $"[{request.Type}] {request.Action} - {request.Summary}";
            txt_meta.Text = $"请求人: {request.RequesterName}    请求时间: {request.CreatedAt:yyyy-MM-dd HH:mm}    状态: {request.Status}"
                + (request.ReviewerName != null ? $"    审核人: {request.ReviewerName} ({request.ReviewedAt:yyyy-MM-dd HH:mm})" : "")
                + (request.ReviewComment != null ? $"    审核意见: {request.ReviewComment}" : "");
            txt_payload.Text = PrettyJson(request.PayloadJson);

            // 非待审核状态不允许再审核
            bool pending = request.Status == ReviewService.StatusPending;
            btn_approve.IsEnabled = pending;
            btn_reject.IsEnabled = pending;
            txt_comment.IsEnabled = pending;
            if (!pending && request.ReviewComment != null)
            {
                txt_comment.Text = request.ReviewComment;
            }
        }

        /// <summary>
        /// 将 payload JSON 格式化输出，失败时原样返回
        /// </summary>
        private static string PrettyJson(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return "(无)";
            }
            try
            {
                return JToken.Parse(json).ToString(Formatting.Indented);
            }
            catch
            {
                return json;
            }
        }

        /* ###############################  事件函数  ################################ */

        private void Btn_Approve_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("确认通过该请求并应用更改？", "审核确认",
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            {
                return;
            }
            string error = _reviewService.Approve(_request.Id, _auth.CurrentOperatorName, txt_comment.Text?.Trim());
            if (error != null)
            {
                _ = MessageBox.Show(error, "审核失败");
                return;
            }
            _ = MessageBox.Show("已通过并应用更改", "审核完成");
            DialogResult = true;
        }

        private void Btn_Reject_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_comment.Text))
            {
                _ = MessageBox.Show("驳回时请填写审核意见", "提示");
                return;
            }
            if (MessageBox.Show("确认驳回该请求？", "审核确认",
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            {
                return;
            }
            string error = _reviewService.Reject(_request.Id, _auth.CurrentOperatorName, txt_comment.Text?.Trim());
            if (error != null)
            {
                _ = MessageBox.Show(error, "审核失败");
                return;
            }
            _ = MessageBox.Show("已驳回", "审核完成");
            DialogResult = true;
        }

        private void Btn_Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}

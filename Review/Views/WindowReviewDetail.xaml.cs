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
            txt_meta.Text = string.Format(LanguageService.Get("ReviewDetail_Meta"), request.RequesterName, request.CreatedAt.ToString("yyyy-MM-dd HH:mm"), request.Status)
                + (request.ReviewerName != null ? string.Format(LanguageService.Get("ReviewDetail_Reviewer"), request.ReviewerName, request.ReviewedAt?.ToString("yyyy-MM-dd HH:mm") ?? "") : "")
                + (request.ReviewComment != null ? string.Format(LanguageService.Get("ReviewDetail_Comment"), request.ReviewComment) : "");
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
            if (MessageBox.Show(LocalizationHelper.Get("Msg_ConfirmApprove"), LanguageService.Get("Cap_ReviewConfirm"),
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            {
                return;
            }
            string error = _reviewService.Approve(_request.Id, _auth.CurrentOperatorName, txt_comment.Text?.Trim());
            if (error != null)
            {
                _ = MessageBox.Show(error, LanguageService.Get("Cap_ReviewFailed"));
                return;
            }
            _ = MessageBox.Show(LocalizationHelper.Get("Msg_Approved"), LanguageService.Get("Cap_ReviewComplete"));
            DialogResult = true;
        }

        private void Btn_Reject_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_comment.Text))
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillRejectReason"), LanguageService.Get("Cap_Info"));
                return;
            }
            if (MessageBox.Show(LocalizationHelper.Get("Msg_ConfirmReject"), LanguageService.Get("Cap_ReviewConfirm"),
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            {
                return;
            }
            string error = _reviewService.Reject(_request.Id, _auth.CurrentOperatorName, txt_comment.Text?.Trim());
            if (error != null)
            {
                _ = MessageBox.Show(error, LanguageService.Get("Cap_ReviewFailed"));
                return;
            }
            _ = MessageBox.Show(LocalizationHelper.Get("Msg_Rejected"), LanguageService.Get("Cap_ReviewComplete"));
            DialogResult = true;
        }

        private void Btn_Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}

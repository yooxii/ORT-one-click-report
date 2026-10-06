using Microsoft.Extensions.DependencyInjection;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using ORT一键报告.Models;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace ORT一键报告.Review.Views
{
    /// <summary>
    /// WindowReviewDetail.xaml 的交互逻辑：请求详情查看与审核（通过/驳回）。
    /// 界面按「头部摘要 → 请求/审核信息 → 更改内容字段表 → 原始数据(折叠) → 审核意见」组织，
    /// 只展示有意义的字段（审计字段、界面用计算字段、空值字段一律不显示）。
    /// </summary>
    public partial class WindowReviewDetail : Window
    {
        private readonly ReviewRequest _request;
        private readonly ReviewService _reviewService;
        private readonly AuthService _auth;

        /// <summary>字段表的行（Label 为本地化后的字段名，Value 为提交值，Changed 表示与库中现值不同）</summary>
        public class FieldRow
        {
            public string Label { get; set; }
            public string Value { get; set; }
            /// <summary>库中原值；仅在 Changed 为 true 时有意义，用于显示「旧值 → 新值」</summary>
            public string OldValue { get; set; }
            public bool Changed { get; set; }
            /// <summary>差异说明：有旧值时显示「旧值 → 新值」，新增字段显示「（原来为空）」</summary>
            public string DiffText { get; set; }
            public bool HasDiff { get; set; }
        }

        /// <summary>不展示的字段：主键、审计字段、以及界面用计算字段（对应 [JsonIgnore]/[IsIgnore]）</summary>
        private static readonly HashSet<string> HiddenFields = new(StringComparer.OrdinalIgnoreCase)
        {
            "Id", "CreatedBy", "CreatedAt", "UpdatedBy", "UpdatedAt",
            "StatusKind", "HasReportLink"
        };

        public WindowReviewDetail(ReviewRequest request)
        {
            InitializeComponent();
            _request = request;
            _reviewService = App.ServiceProvider.GetRequiredService<ReviewService>();
            _auth = App.ServiceProvider.GetRequiredService<AuthService>();

            FillHeader(request);
            FillRequestInfo(request);
            FillChangedFields(request);
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
            // 已审结的请求把结论显示在按钮行左侧，避免审核人重复核对
            if (!pending && request.ReviewerName != null)
            {
                txt_decided.Text = string.Format(LanguageService.Get("ReviewDetail_Reviewer"), request.ReviewerName,
                    request.ReviewedAt?.ToString("yyyy-MM-dd HH:mm") ?? "");
            }
            lbl_reviewed.Visibility = request.ReviewerName != null ? Visibility.Visible : Visibility.Collapsed;
        }

        /* ###############################  填充  ################################ */

        /// <summary>头部：操作徽标 + 标题 + 目标记录 + 状态徽标（配色随状态变化）</summary>
        private void FillHeader(ReviewRequest request)
        {
            txt_action.Text = request.Action ?? "";
            badge_action.Background = ActionBrush(request.Action);
            txt_title.Text = $"{request.Type} - {request.Summary}";
            txt_target.Text = request.TargetId != null
                ? string.Format(LanguageService.Get("ReviewDetail_Target"), request.TargetId.Value)
                : "";
            txt_target.Visibility = string.IsNullOrEmpty(txt_target.Text) ? Visibility.Collapsed : Visibility.Visible;

            txt_status.Text = request.Status ?? "";
            badge_status.Background = StatusBackground(request.Status);
        }

        /// <summary>请求/审核信息：两列排布，取代原先拼成一长串的 meta 文本</summary>
        private void FillRequestInfo(ReviewRequest request)
        {
            txt_requester.Text = request.RequesterName ?? "-";
            txt_requestTime.Text = request.CreatedAt.ToString("yyyy-MM-dd HH:mm");
            txt_assignee.Text = request.AssigneeName ?? "-";
            txt_reviewed.Text = request.ReviewerName == null
                ? "-"
                : $"{request.ReviewerName} ({request.ReviewedAt?.ToString("yyyy-MM-dd HH:mm") ?? "-"})";
        }

        /// <summary>
        /// 更改内容：把 payload 反序列化为对象后逐字段列出，只保留有值的字段；
        /// 编辑类请求与库中现值比对，不同的字段高亮。
        /// </summary>
        private void FillChangedFields(ReviewRequest request)
        {
            List<FieldRow> rows = [];
            if (!string.IsNullOrWhiteSpace(request.PayloadJson))
            {
                try
                {
                    if (request.Type == "领退表单")
                    {
                        rows = BuildRows(JsonConvert.DeserializeObject<Requisition>(request.PayloadJson), request);
                    }
                    else
                    {
                        rows = BuildRows(JsonConvert.DeserializeObject<Plan>(request.PayloadJson), request);
                    }
                }
                catch
                {
                    // payload 解析失败时不显示字段表，原始数据里仍可查看
                    rows = [];
                }
            }

            ic_fields.ItemsSource = rows;
            bool empty = rows.Count == 0;
            txt_noFields.Visibility = empty ? Visibility.Visible : Visibility.Collapsed;
            sv_fields.Visibility = empty ? Visibility.Collapsed : Visibility.Visible;
            exp_raw.IsExpanded = empty;   // 没有可读字段时直接展开原始数据
        }

        /// <summary>把对象序列化成 JObject 再逐字段取值，避免为两个模型各写一遍映射</summary>
        private List<FieldRow> BuildRows(object payload, ReviewRequest request)
        {
            if (payload == null)
            {
                return [];
            }
            JObject current = request.Action == "编辑" ? LoadCurrent(request) : null;
            JObject payloadObj = JObject.FromObject(payload);
            List<FieldRow> rows = [];
            foreach (JProperty property in payloadObj.Properties())
            {
                if (HiddenFields.Contains(property.Name))
                {
                    continue;
                }
                string value = FormatValue(property.Value);
                if (value.Length == 0)
                {
                    continue;   // 空值字段不占位
                }
                string oldValue = null;
                bool changed = false;
                if (current != null && current.TryGetValue(property.Name, out JToken old))
                {
                    oldValue = FormatValue(old);
                    changed = oldValue != value;
                }
                rows.Add(new FieldRow
                {
                    Label = FieldLabel(property.Name),
                    Value = value,
                    OldValue = changed ? oldValue : null,
                    Changed = changed,
                    HasDiff = changed,
                    // 原值为空时说明是「原来没填、这次补上」，用专门文案而不是空箭头
                    DiffText = changed
                        ? string.Format(
                            LanguageService.Get(oldValue.Length == 0 ? "ReviewDetail_OldEmpty" : "ReviewDetail_OldToNew"),
                            oldValue, value)
                        : null
                });
            }
            return rows;
        }

        /// <summary>读取目标记录的当前值，用于标记差异字段（读不到就不做高亮）</summary>
        private JObject LoadCurrent(ReviewRequest request)
        {
            if (request.TargetId == null)
            {
                return null;
            }
            try
            {
                DatabaseService db = App.ServiceProvider.GetRequiredService<DatabaseService>();
                object entity = request.Type == "领退表单"
                    ? db.FreeSql.Select<Requisition>().Where(r => r.Id == request.TargetId.Value).First()
                    : db.FreeSql.Select<Plan>().Where(p => p.Id == request.TargetId.Value).First();
                return entity == null ? null : JObject.FromObject(entity);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>字段取值格式化：日期只留日期部分，长文本截断，其余原样</summary>
        private static string FormatValue(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null || token.Type == JTokenType.Undefined)
            {
                return "";
            }
            if (token.Type == JTokenType.Date)
            {
                return token.Value<DateTime>().ToString("yyyy-MM-dd");
            }
            string text = token.Type == JTokenType.String ? token.Value<string>() : token.ToString(Formatting.None);
            text = text?.Trim() ?? "";
            if (text.Length == 0 || text is "(无)" or "[]" or "{}")
            {
                return "";
            }
            // 序列号清单等超长文本只给预览，完整内容在「原始数据」里查看
            return text.Length > MaxValueLength ? text.Substring(0, MaxValueLength) + "…" : text;
        }

        /// <summary>字段表中单个值的最大长度（超出截断，完整值见原始数据）</summary>
        private const int MaxValueLength = 400;

        /// <summary>字段名 → 本地化标签（键缺失时 Get 原样返回键名，所以这里回退到字段名）</summary>
        private static string FieldLabel(string name)
        {
            string key = FieldLabelKeys.TryGetValue(name, out string value) ? value : null;
            if (key == null)
            {
                return name;
            }
            string text = LanguageService.Get(key);
            return text == key ? name : text;
        }

        /// <summary>模型字段名到资源键的映射（沿用各编辑窗口已有的标签文案）</summary>
        private static readonly Dictionary<string, string> FieldLabelKeys = new(StringComparer.OrdinalIgnoreCase)
        {
            // 计划表
            ["JobNo"] = "PlanEdit_WorkNumber",
            ["Product"] = "Plans_ProductType",
            ["Customer"] = "PlanEdit_Customer",
            ["ModelName"] = "Plans_ModelName",
            ["Stage"] = "Plans_Stage",
            ["TestItem"] = "Plans_TestItem",
            ["SampleSize"] = "PlanEdit_SampleQty",
            ["TestPeriod"] = "Admin_TestTime",
            ["Owner"] = "Plans_Owner",
            ["StartDate"] = "Plans_StartDate",
            ["EndDate"] = "PlanEdit_EndDate",
            ["Status"] = "Plans_Status",
            ["ReportStatus"] = "Plans_ReportStatus",
            ["UnitReturnDate"] = "Plans_ReturnDate",
            ["Remark"] = "Admin_Remark",
            // 领退表
            ["RequisitionDate"] = "Plans_RequisitionDateTC",
            ["RequisitionNo"] = "Plans_RequisitionDocNo",
            ["OutQty"] = "Plans_IssueQty",
            ["Disposition"] = "Plans_Disposition",
            ["SN"] = "Plans_SN",
            ["SnFilePath"] = "Plans_SNFile",
            ["Rev"] = "Plans_REV",
            ["WorkOrder"] = "Plans_WorkOrder",
            ["DC"] = "Plans_DC",
            ["LineNo"] = "Plans_Line",
            ["ReturnRtOrder"] = "Plans_ReturnRTOrder",
            ["ReturnQty"] = "Plans_ReturnQty",
            ["ReturnDate"] = "Plans_ReturnDate",
            ["StockInNo"] = "Plans_StockInNo",
            ["StockInQty"] = "Plans_StockInQty",
            ["StockInDate"] = "Plans_StockInDate",
            ["ScrapNo"] = "Plans_ScrapNo",
            ["ScrapQty"] = "Plans_ScrapQty",
            ["ScrapDate"] = "Plans_ScrapDate",
            ["ScrapSnText"] = "Plans_ScrapSn",
            ["ScrapSnFilePath"] = "Plans_ScrapSnFile"
        };

        /* ###############################  配色  ################################ */

        /// <summary>操作徽标底色：新增/编辑/删除/报废分别取主题里的成功/主题色/错误/警告色</summary>
        private static Brush ActionBrush(string action) => action switch
        {
            "新增" => Res("StatusOkBgBrush", Color.FromRgb(0xE8, 0xF5, 0xE9)),
            "编辑" => Res("StatusOngoingBgBrush", Color.FromRgb(0xE3, 0xF2, 0xFD)),
            "删除" => Res("StatusErrorBgBrush", Color.FromRgb(0xFF, 0xEB, 0xEE)),
            "报废" => Res("StatusWarnBgBrush", Color.FromRgb(0xFF, 0xF3, 0xE0)),
            _ => Res("HeaderBackgroundBrush", Color.FromRgb(0xF5, 0xF5, 0xF5))
        };

        /// <summary>状态徽标底色：待审核/已通过/已驳回</summary>
        private static Brush StatusBackground(string status) => status switch
        {
            ReviewService.StatusApproved => Res("StatusOkBgBrush", Color.FromRgb(0xE8, 0xF5, 0xE9)),
            ReviewService.StatusRejected => Res("StatusErrorBgBrush", Color.FromRgb(0xFF, 0xEB, 0xEE)),
            _ => Res("StatusPendingBgBrush", Color.FromRgb(0xFF, 0xF3, 0xE0))
        };

        /// <summary>取主题画刷；主题未初始化时（如单测/设计器）退回给定色值</summary>
        private static Brush Res(string key, Color fallback)
            => Application.Current?.TryFindResource(key) as Brush ?? new SolidColorBrush(fallback);

        /* ###############################  其他  ################################ */

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
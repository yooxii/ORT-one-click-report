using NLog;
using ORT一键报告.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 邮件业务触发：
    /// 警告类——计划表临近/超过测试结束日期且仍未生成报告 → 通知负责人；
    /// 通知类——审核请求提交/审核结果/密码变更 → 通知审核员或相关人员。
    /// 所有发送均容错：邮件失败只记日志，绝不影响主业务流程。
    /// </summary>
    public class MailNotifier
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;
        private readonly AppSettingsService _settings;
        private readonly MailService _mail;

        public MailNotifier(DatabaseService db, AppSettingsService settings, MailService mail)
        {
            _db = db;
            _settings = settings;
            _mail = mail;
        }

        /* ###############################  警告类：临近结束日期  ################################ */

        /// <summary>
        /// 检查计划表并发送"临近结束日期且未完成"警告。
        /// 条件：有结束日期；结束日期在 今天+N 天 内（可选含已逾期）；
        ///       完成状况 ≠ Close（未完成，空值视为未完成）。
        /// 返回（发送成功数, 跳过数, 失败数）。
        /// </summary>
        public (int Sent, int Skipped, int Failed) CheckPlanDeadlines(DateTime? today = null)
        {
            int sent = 0, skipped = 0, failed = 0;
            MailSettings mail = _settings.Settings.Mail;
            if (!mail.Enabled || !mail.WarningEnabled)
            {
                return (0, 0, 0);
            }
            DateTime date = (today ?? DateTime.Now).Date;
            DateTime limit = date.AddDays(Math.Max(0, mail.WarningDaysBefore));
            List<Plan> plans;
            try
            {
                plans = _db.FreeSql.Select<Plan>().Where(p => p.EndDate != null && p.EndDate <= limit).ToList();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "读取计划表失败，跳过结束日期提醒");
                return (0, 0, 0);
            }
            foreach (Plan plan in plans)
            {
                DateTime end = plan.EndDate.Value.Date;
                if (end < date && !mail.WarningIncludeOverdue)
                {
                    continue; // 已逾期且不提醒逾期
                }
                if (IsCompleted(plan.Status))
                {
                    continue; // 完成状况为 Close → 视为已完成，不再提醒
                }
                int daysLeft = (int)(end - date).TotalDays;
                Dictionary<string, object> data = new()
                {
                    ["JobNo"] = plan.JobNo,
                    ["ModelName"] = plan.ModelName,
                    ["TestItem"] = plan.TestItem,
                    ["TestPeriod"] = plan.TestPeriod,
                    ["Stage"] = plan.Stage,
                    ["Customer"] = plan.Customer,
                    ["Product"] = plan.Product,
                    ["Owner"] = plan.Owner,
                    ["StartDate"] = plan.StartDate,
                    ["EndDate"] = plan.EndDate,
                    ["DaysLeft"] = daysLeft,
                    ["Overdue"] = daysLeft < 0,
                    ["Status"] = plan.Status,
                    ["Remark"] = plan.Remark,
                    ["Content"] = string.Format(LanguageService.Get("Mail_WarningContentFormat"),
                        plan.JobNo ?? "-", plan.ModelName ?? "-", end.ToString("yyyy/M/d"), daysLeft, plan.Status ?? "-"),
                };
                MailSendResult result = _mail.Send(MailKind.Warning, MailService.SplitList(plan.Owner),
                    data, "Plan", plan.JobNo, false, AdminCc(MailKind.Warning));
                if (result.Success)
                {
                    sent++;
                }
                else if (result.Skipped)
                {
                    skipped++;
                }
                else
                {
                    failed++;
                }
            }
            if (sent > 0 || failed > 0)
            {
                _logger.Info($"计划结束日期提醒: 成功{sent} 跳过{skipped} 失败{failed}");
            }
            return (sent, skipped, failed);
        }

        /// <summary>
        /// 完成状况是否为已结案（Close；空值视为未完成）
        /// </summary>
        public static bool IsCompleted(string status)
            => string.Equals(status?.Trim(), "Close", StringComparison.OrdinalIgnoreCase);

        /// <summary>
        /// 该类型需要抄送的管理员邮箱（设置中按类型开关；未开启返回 null）
        /// </summary>
        private List<string> AdminCc(string kind)
            => _settings.Settings.Mail.ShouldCcAdmins(MailKind.Find(kind))
                ? _mail.GetRoleEmails(UserRole.Administrator)
                : null;

        /// <summary>
        /// 后台执行结束日期检查（不阻塞界面）
        /// </summary>
        public void CheckPlanDeadlinesInBackground()
        {
            MailSettings mail = _settings.Settings.Mail;
            if (!mail.Enabled || !mail.WarningEnabled)
            {
                return;
            }
            _ = Task.Run(() =>
            {
                try
                {
                    CheckPlanDeadlines();
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "后台检查计划结束日期失败");
                }
            });
        }

        /* ###############################  通知类  ################################ */

        /// <summary>
        /// 审核请求提交 → 只通知"当前待审核人"（提交时指派；未指派时退回全部审核员），按设置抄送管理员
        /// </summary>
        public void NotifyReviewSubmitted(ReviewRequest request)
        {
            if (request == null)
            {
                return;
            }
            RunInBackground(() =>
            {
                List<string> recipients = string.IsNullOrWhiteSpace(request.AssigneeName)
                    ? _mail.GetRoleEmails(UserRole.Reviewer)
                    : [request.AssigneeName];
                Dictionary<string, object> data = new()
                {
                    ["Title"] = LanguageService.Get("Mail_ReviewSubmittedTitle"),
                    ["Category"] = LanguageService.Get(MailKind.Find(MailKind.Notice).NameKey),
                    ["RequestType"] = request.Type,
                    ["Action"] = request.Action,
                    ["Summary"] = request.Summary,
                    ["Requester"] = request.RequesterName,
                    ["Assignee"] = request.AssigneeName,
                    ["Status"] = request.Status,
                    ["Comment"] = request.ReviewComment,
                    ["Reviewer"] = request.ReviewerName,
                    ["Content"] = string.Format(LanguageService.Get("Mail_ReviewSubmittedContentFormat"),
                        request.RequesterName ?? "-", request.Summary ?? "-"),
                };
                _mail.Send(MailKind.Notice, recipients, data, "Review", "Submit#" + request.Id, true,
                    AdminCc(MailKind.Notice));
            });
        }

        /// <summary>
        /// 审核通过/驳回 → 通知请求人，按设置抄送管理员
        /// </summary>
        public void NotifyReviewDecided(ReviewRequest request, bool approved)
        {
            if (request == null)
            {
                return;
            }
            RunInBackground(() =>
            {
                Dictionary<string, object> data = new()
                {
                    ["Title"] = LanguageService.Get(approved ? "Mail_ReviewApprovedTitle" : "Mail_ReviewRejectedTitle"),
                    ["Category"] = LanguageService.Get(MailKind.Find(MailKind.Notice).NameKey),
                    ["RequestType"] = request.Type,
                    ["Action"] = request.Action,
                    ["Summary"] = request.Summary,
                    ["Requester"] = request.RequesterName,
                    ["Assignee"] = request.AssigneeName,
                    ["Status"] = request.Status,
                    ["Comment"] = request.ReviewComment,
                    ["Reviewer"] = request.ReviewerName,
                    ["Content"] = string.Format(LanguageService.Get("Mail_ReviewDecidedContentFormat"),
                        request.Summary ?? "-", request.Status ?? "-", request.ReviewComment ?? "-"),
                };
                _mail.Send(MailKind.Notice, [request.RequesterName], data, "Review", "Decide#" + request.Id, true,
                    AdminCc(MailKind.Notice));
            });
        }

        /// <summary>
        /// 密码变更 → 通知本人
        /// </summary>
        public void NotifyPasswordChanged(string username, string displayName, string changedBy)
        {
            RunInBackground(() =>
            {
                Dictionary<string, object> data = new()
                {
                    ["Title"] = LanguageService.Get("Mail_PasswordChangedTitle"),
                    ["Category"] = LanguageService.Get(MailKind.Find(MailKind.Notice).NameKey),
                    ["Username"] = username,
                    ["DisplayName"] = displayName,
                    ["ChangedBy"] = changedBy,
                    ["Status"] = LanguageService.Get("Mail_PasswordChangedStatus"),
                    ["Content"] = string.Format(LanguageService.Get("Mail_PasswordChangedContentFormat"), username ?? "-"),
                };
                _mail.Send(MailKind.Notice, [username], data, "User", "Password#" + username, true);
            });
        }

        /// <summary>
        /// 登录时提示补充邮箱 → 提供"稍后可在设置中完善"的说明邮件（可选）。
        /// 目前仅用于手工发送测试，保留扩展点。
        /// </summary>
        public void NotifyEmailMissing(string username, string displayName)
        {
            RunInBackground(() =>
            {
                Dictionary<string, object> data = new()
                {
                    ["Title"] = LanguageService.Get("Mail_EmailMissingTitle"),
                    ["Category"] = LanguageService.Get(MailKind.Find(MailKind.Notice).NameKey),
                    ["Username"] = username,
                    ["DisplayName"] = displayName,
                    ["Content"] = string.Format(LanguageService.Get("Mail_EmailMissingContentFormat"), displayName ?? username ?? "-"),
                };
                _mail.Send(MailKind.Notice, [username], data, "User", "EmailMissing#" + username, true);
            });
        }

        /// <summary>
        /// 后台执行：邮件失败只记日志，不影响业务
        /// </summary>
        private void RunInBackground(Action action)
        {
            MailSettings mail = _settings.Settings.Mail;
            if (!mail.Enabled || !mail.NoticeEnabled)
            {
                return;
            }
            _ = Task.Run(() =>
            {
                try
                {
                    action();
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "发送通知邮件失败");
                }
            });
        }
    }
}

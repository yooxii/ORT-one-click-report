using NLog;
using ORT一键报告.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 邮件发送服务：按设置组装 SMTP 连接、渲染模板、写入发送记录并做重复发送抑制。
    /// 业务触发（谁在什么情况下收信）见 <see cref="MailNotifier"/>。
    /// </summary>
    public class MailService
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;
        private readonly AppSettingsService _settings;

        public MailService(DatabaseService db, AppSettingsService settings)
        {
            _db = db;
            _settings = settings;
        }

        /// <summary>邮件设置（实时读取，修改设置后立即生效）</summary>
        public MailSettings Settings => _settings.Settings.Mail;

        /// <summary>是否已具备发送条件；不满足时 reason 说明原因</summary>
        public bool IsReady(out string reason)
        {
            MailSettings mail = Settings;
            if (!mail.Enabled)
            {
                reason = LanguageService.Get("Msg_MailDisabled");
                return false;
            }
            if (string.IsNullOrWhiteSpace(mail.Host))
            {
                reason = LanguageService.Get("Msg_MailNoHost");
                return false;
            }
            if (!IsValidAddress(mail.FromAddress))
            {
                reason = LanguageService.Get("Msg_MailNoFrom");
                return false;
            }
            reason = null;
            return true;
        }

        /// <summary>
        /// 发送指定类型的邮件；返回结果（不抛异常，失败原因写日志与 mail_logs）
        /// </summary>
        /// <param name="kind">邮件类型代码（<see cref="MailKind"/>）</param>
        /// <param name="recipients">收件人（姓名或邮箱均可为空，空则跳过）</param>
        /// <param name="variables">模板变量（自动补充公共变量）</param>
        /// <param name="refType">关联对象类型（Plan/Review/User/Settings）</param>
        /// <param name="refKey">关联对象键（用于重复发送抑制）</param>
        /// <param name="bypassDedupe">忽略重复发送抑制（手工测试用）</param>
        /// <param name="cc">本次额外抄送（如按设置抄送管理员）</param>
        public MailSendResult Send(string kind, IEnumerable<string> recipients,
            IDictionary<string, object> variables, string refType = null, string refKey = null,
            bool bypassDedupe = false, IEnumerable<string> cc = null)
        {
            MailTypeDefinition type = MailKind.Find(kind);
            if (type == null)
            {
                return MailSendResult.Skip($"未知邮件类型: {kind}");
            }
            if (!IsReady(out string reason))
            {
                return MailSendResult.Skip(reason);
            }
            if (!bypassDedupe && WasSentRecently(type.Code, refType, refKey, Settings.DedupeDays))
            {
                return MailSendResult.Skip(LanguageService.Get("Msg_MailDeduped"));
            }
            List<string> to = NormalizeRecipients(recipients);
            if (to.Count == 0)
            {
                return MailSendResult.Skip(LanguageService.Get("Msg_MailNoRecipient"));
            }

            Dictionary<string, object> data = MailTemplate.BuildBaseVariables();
            foreach (KeyValuePair<string, object> pair in variables ?? new Dictionary<string, object>())
            {
                data[pair.Key] = pair.Value;
            }
            // 收件人与发件人也可在模板中使用
            data["Recipients"] = string.Join("; ", to);
            data["Sender"] = Settings.FromName ?? Settings.FromAddress;

            string templateSubject = Settings.GetTemplate(type, true) ?? LanguageService.Get(type.DefaultSubjectKey);
            string templateBody = Settings.GetTemplate(type, false) ?? LanguageService.Get(type.DefaultBodyKey);
            string subject = MailTemplate.Render(templateSubject, data).Replace("\r", " ").Replace("\n", " ").Trim();
            string body = MailTemplate.Render(templateBody, data);

            MailMessage message = new()
            {
                From = Settings.FromAddress.Trim(),
                FromName = Settings.FromName,
                Subject = subject,
                Body = body,
                IsHtml = Settings.BodyIsHtml
            };
            message.To.AddRange(to);
            message.Cc.AddRange(NormalizeRecipients(SplitList(Settings.CcList)));
            foreach (string address in NormalizeRecipients(cc ?? []))
            {
                if (!message.To.Contains(address, StringComparer.OrdinalIgnoreCase)
                    && !message.Cc.Contains(address, StringComparer.OrdinalIgnoreCase))
                {
                    message.Cc.Add(address);
                }
            }

            try
            {
                new SmtpMailClient(Settings).Send(message);
                WriteLog(type.Code, refType, refKey, message, true, null);
                return MailSendResult.Ok(subject, string.Join(";", to));
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"发送邮件失败({type.Code}): {subject}");
                WriteLog(type.Code, refType, refKey, message, false, ex.Message);
                return MailSendResult.Fail(ex.Message, subject, string.Join(";", to));
            }
        }

        /// <summary>
        /// 发送测试邮件（忽略重复抑制）
        /// </summary>
        public MailSendResult SendTest(string recipient)
        {
            MailTypeDefinition type = MailKind.Find(MailKind.Notice);
            Dictionary<string, object> data = new()
            {
                ["Title"] = LanguageService.Get("Mail_TestSubject"),
                ["Content"] = LanguageService.Get("Mail_TestBody"),
                ["Category"] = LanguageService.Get(type.NameKey),
                ["Username"] = Settings.Username,
                ["DisplayName"] = Settings.FromName,
            };
            return Send(MailKind.Notice, [recipient], data, "Settings", "TestMail", true);
        }

        /// <summary>
        /// 指定对象近期是否已发送过同类邮件（重复发送抑制）
        /// </summary>
        public bool WasSentRecently(string kind, string refType, string refKey, int days)
        {
            if (string.IsNullOrWhiteSpace(refKey) || days <= 0)
            {
                return false;
            }
            try
            {
                DateTime since = DateTime.Now.AddDays(-days);
                return _db.FreeSql.Select<MailLog>()
                    .Where(l => l.Kind == kind && l.RefType == refType && l.RefKey == refKey
                                && l.Success && l.CreatedAt >= since)
                    .Any();
            }
            catch (Exception ex)
            {
                _logger.Warn($"查询邮件发送记录失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 最近发送记录（设置界面排查用）
        /// </summary>
        public List<MailLog> RecentLogs(int count = 20)
        {
            try
            {
                return _db.FreeSql.Select<MailLog>().OrderByDescending(l => l.Id).Limit(count).ToList();
            }
            catch (Exception ex)
            {
                _logger.Warn($"读取邮件发送记录失败: {ex.Message}");
                return [];
            }
        }

        /// <summary>
        /// 邮箱格式简单校验
        /// </summary>
        public static bool IsValidAddress(string address)
            => !string.IsNullOrWhiteSpace(address)
               && System.Text.RegularExpressions.Regex.IsMatch(address.Trim(), @"^[^@\s]+@[^@\s]+\.[^@\s]+$");

        /// <summary>
        /// 收件人规范化：支持"姓名"或"邮箱"输入——姓名会去用户表解析邮箱；过滤非法地址并去重
        /// </summary>
        public List<string> NormalizeRecipients(IEnumerable<string> recipients)
        {
            List<string> result = [];
            foreach (string raw in recipients ?? [])
            {
                foreach (string item in SplitList(raw))
                {
                    string address = item;
                    if (!IsValidAddress(address))
                    {
                        address = ResolveUserEmail(item);
                    }
                    if (IsValidAddress(address) && !result.Contains(address, StringComparer.OrdinalIgnoreCase))
                    {
                        result.Add(address.Trim());
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 按显示名/登录名解析用户邮箱（找不到返回 null）
        /// </summary>
        public string ResolveUserEmail(string nameOrUsername)
        {
            if (string.IsNullOrWhiteSpace(nameOrUsername))
            {
                return null;
            }
            string key = nameOrUsername.Trim();
            List<User> users = _db.FreeSql.Select<User>().ToList();
            User user = users.FirstOrDefault(u => string.Equals(u.DisplayName, key, StringComparison.CurrentCultureIgnoreCase))
                ?? users.FirstOrDefault(u => string.Equals(u.Username, key, StringComparison.OrdinalIgnoreCase));
            return string.IsNullOrWhiteSpace(user?.Email) ? null : user.Email.Trim();
        }

        /// <summary>
        /// 按角色取邮箱列表（如审核员）
        /// </summary>
        public List<string> GetRoleEmails(UserRole role)
        {
            List<User> users = _db.FreeSql.Select<User>().Where(u => u.IsActive).ToList();
            List<long> roleUserIds = _db.FreeSql.Select<UserRoleRow>()
                .Where(r => r.Role == role.ToString()).ToList(r => r.UserId);
            return users.Where(u => roleUserIds.Contains(u.Id) && IsValidAddress(u.Email))
                .Select(u => u.Email.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private void WriteLog(string kind, string refType, string refKey, MailMessage message, bool success, string error)
        {
            try
            {
                _db.FreeSql.Insert(new MailLog
                {
                    Kind = kind,
                    RefType = refType,
                    RefKey = refKey,
                    Recipients = string.Join(";", message.To.Concat(message.Cc)),
                    Subject = message.Subject,
                    Success = success,
                    Error = error,
                    CreatedAt = DateTime.Now
                }).ExecuteAffrows();
            }
            catch (Exception ex)
            {
                _logger.Warn($"写入邮件发送记录失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 拆分收件人文本（逗号/分号/中文分号/换行）
        /// </summary>
        public static List<string> SplitList(string text)
            => string.IsNullOrWhiteSpace(text)
                ? []
                : text.Split([',', ';', '；', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim())
                    .Where(s => s.Length > 0)
                    .ToList();
    }
}

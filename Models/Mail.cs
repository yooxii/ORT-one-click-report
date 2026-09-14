using FreeSql.DataAnnotations;
using System;

namespace ORT一键报告.Models
{
    /// <summary>
    /// 邮件类型定义（可扩展）：新增类型只需在 <see cref="MailKind.All"/> 增加一项并补充对应资源键。
    /// 模板内容按 Code 存于设置中（mail.template.&lt;Code&gt;.subject / .body）。
    /// </summary>
    public class MailTypeDefinition
    {
        /// <summary>类型代码（Notice/Warning…）</summary>
        public string Code { get; }

        /// <summary>类型显示名资源键</summary>
        public string NameKey { get; }

        /// <summary>默认主题模板资源键</summary>
        public string DefaultSubjectKey { get; }

        /// <summary>默认正文模板资源键</summary>
        public string DefaultBodyKey { get; }

        /// <summary>模板可用变量说明资源键</summary>
        public string VariablesKey { get; }

        public MailTypeDefinition(string code, string nameKey, string defaultSubjectKey, string defaultBodyKey, string variablesKey)
        {
            Code = code;
            NameKey = nameKey;
            DefaultSubjectKey = defaultSubjectKey;
            DefaultBodyKey = defaultBodyKey;
            VariablesKey = variablesKey;
        }

        public string SubjectSettingKey => "mail.template." + Code + ".subject";

        public string BodySettingKey => "mail.template." + Code + ".body";
    }

    /// <summary>
    /// 邮件类型：通知类 / 警告类（后续可继续增加）
    /// </summary>
    public static class MailKind
    {
        /// <summary>通知类（审核通知、密码变更等）</summary>
        public const string Notice = "Notice";

        /// <summary>警告类（临近测试结束日期未完成报告等）</summary>
        public const string Warning = "Warning";

        /// <summary>
        /// 全部类型（顺序即界面展示顺序）
        /// </summary>
        public static readonly MailTypeDefinition[] All =
        [
            new MailTypeDefinition(Notice, "Mail_Kind_Notice", "Mail_Default_Notice_Subject", "Mail_Default_Notice_Body", "Mail_Vars_Notice"),
            new MailTypeDefinition(Warning, "Mail_Kind_Warning", "Mail_Default_Warning_Subject", "Mail_Default_Warning_Body", "Mail_Vars_Warning"),
        ];

        public static MailTypeDefinition Find(string code)
        {
            foreach (MailTypeDefinition type in All)
            {
                if (string.Equals(type.Code, code, StringComparison.OrdinalIgnoreCase))
                {
                    return type;
                }
            }
            return null;
        }
    }

    /// <summary>
    /// SMTP 安全方式选项（设置界面下拉框绑定用）
    /// </summary>
    public class MailSecurityOption
    {
        /// <summary>Options: None / StartTls / Ssl</summary>
        public string Code { get; }

        /// <summary>显示名（本地化）</summary>
        public string DisplayName { get; }

        public MailSecurityOption(string code, string displayName)
        {
            Code = code;
            DisplayName = displayName;
        }
    }

    /// <summary>
    /// 邮件发送记录（mail_logs 表）：用于去重（同一对象同类邮件间隔）与排查
    /// </summary>
    [Table(Name = "mail_logs")]
    public class MailLog
    {
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        /// <summary>邮件类型代码（Notice/Warning）</summary>
        [Column(StringLength = 16, IsNullable = false)]
        public string Kind { get; set; }

        /// <summary>关联对象类型（Plan/Review/User/Settings）</summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string RefType { get; set; }

        /// <summary>关联对象键（计划=工作編號，审核=请求Id，用户=登录名）</summary>
        [Column(StringLength = 128, IsNullable = true)]
        public string RefKey { get; set; }

        /// <summary>收件人（分号分隔）</summary>
        [Column(StringLength = 1024, IsNullable = true)]
        public string Recipients { get; set; }

        [Column(StringLength = 512, IsNullable = true)]
        public string Subject { get; set; }

        public bool Success { get; set; }

        [Column(StringLength = 1024, IsNullable = true)]
        public string Error { get; set; }

        public DateTime CreatedAt { get; set; } = DateTime.Now;
    }

    /// <summary>
    /// 邮件发送结果
    /// </summary>
    public class MailSendResult
    {
        /// <summary>是否成功发出</summary>
        public bool Success { get; set; }

        /// <summary>是否因未启用/未配置/去重而跳过</summary>
        public bool Skipped { get; set; }

        /// <summary>跳过或失败原因</summary>
        public string Message { get; set; }

        public string Subject { get; set; }

        public string Recipients { get; set; }

        public static MailSendResult Ok(string subject, string recipients)
            => new() { Success = true, Subject = subject, Recipients = recipients };

        public static MailSendResult Skip(string reason)
            => new() { Skipped = true, Message = reason };

        public static MailSendResult Fail(string error, string subject = null, string recipients = null)
            => new() { Success = false, Message = error, Subject = subject, Recipients = recipients };
    }
}

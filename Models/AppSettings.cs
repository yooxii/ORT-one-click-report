using FreeSql.DataAnnotations;
using System.Collections.Generic;

namespace ORT一键报告.Models
{
    /// <summary>
    /// 设置键值对实体（app_settings 表）：所有设置项以键值对形式保存到数据库。
    /// 注：数据库路径设置项单独保存在程序目录文件中（避免自引用）。
    /// </summary>
    [Table(Name = "app_settings")]
    [Index("uk_app_setting_key", nameof(Key), true)]
    public class AppSetting
    {
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        /// <summary>
        /// 设置键（如 ui.fontFamily / paths.report）
        /// </summary>
        [Column(StringLength = 64, IsNullable = false)]
        public string Key { get; set; }

        /// <summary>
        /// 设置值
        /// </summary>
        [Column(StringLength = 512, IsNullable = true)]
        public string Value { get; set; }
    }

    /// <summary>
    /// 报告链接实体（report_links 表）：记录按 RT 工作编号在报告路径下找到的报告文件夹。
    /// 报告夹结构：文件夹名包含工作编号，内含 Report 子文件夹与一个 Excel 报告概览文件。
    /// </summary>
    [Table(Name = "report_links")]
    [Index("uk_report_link_job", nameof(JobNo), true)]
    public class ReportLink
    {
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        /// <summary>
        /// 工作编号（与计划表对应）
        /// </summary>
        [Column(StringLength = 64, IsNullable = false)]
        public string JobNo { get; set; }

        /// <summary>
        /// Report 子文件夹完整路径（打开报告文件夹目标）
        /// </summary>
        [Column(StringLength = 512, IsNullable = true)]
        public string ReportDir { get; set; }

        /// <summary>
        /// 报告概览 Excel 文件完整路径（与 Report 同级）
        /// </summary>
        [Column(StringLength = 512, IsNullable = true)]
        public string OverviewFile { get; set; }

        /// <summary>
        /// 扫描时间
        /// </summary>
        [Column(IsNullable = true)]
        public System.DateTime? UpdatedAt { get; set; }
    }

    /// <summary>
    /// 应用设置（保存到数据库 app_settings 表；数据库路径单独保存在程序目录）
    /// </summary>
    public class AppSettings
    {
        /// <summary>
        /// UI 设置
        /// </summary>
        public UiSettings UI { get; set; } = new UiSettings();

        /// <summary>
        /// 邮件服务设置
        /// </summary>
        public MailSettings Mail { get; set; } = new MailSettings();

        /// <summary>
        /// 路径设置
        /// </summary>
        public PathSettings Paths { get; set; } = new PathSettings();
    }

    /// <summary>
    /// 邮件服务设置：发件账号（服务器/端口/安全方式/认证）、触发规则与邮件模板。
    /// 密码以 DPAPI（当前 Windows 用户）加密后存库。
    /// </summary>
    public class MailSettings
    {
        /// <summary>总开关</summary>
        public bool Enabled { get; set; }

        /// <summary>通知类启用（审核通知、密码变更等）</summary>
        public bool NoticeEnabled { get; set; } = true;

        /// <summary>警告类启用（临近测试结束日期等）</summary>
        public bool WarningEnabled { get; set; } = true;

        /// <summary>SMTP 服务器</summary>
        public string Host { get; set; }

        /// <summary>端口（常用：25 明文 / 587 STARTTLS / 465 SSL）</summary>
        public int Port { get; set; } = 25;

        /// <summary>安全方式：None=明文，StartTls=显式TLS，Ssl=隐式TLS（465）</summary>
        public string Security { get; set; } = "None";

        /// <summary>忽略服务器证书错误（自签名证书的内网邮件服务器）</summary>
        public bool IgnoreCertErrors { get; set; }

        /// <summary>使用 Windows 集成身份验证（当前登录用户凭据）</summary>
        public bool UseDefaultCredentials { get; set; }

        /// <summary>登录账号</summary>
        public string Username { get; set; }

        /// <summary>登录密码（内存明文；落库前 DPAPI 加密）</summary>
        public string Password { get; set; }

        /// <summary>发件地址</summary>
        public string FromAddress { get; set; }

        /// <summary>发件人显示名</summary>
        public string FromName { get; set; }

        /// <summary>固定抄送（分号或逗号分隔，可留空）</summary>
        public string CcList { get; set; }

        /// <summary>正文按 HTML 发送</summary>
        public bool BodyIsHtml { get; set; }

        /// <summary>发送超时（秒）</summary>
        public int TimeoutSeconds { get; set; } = 30;

        /// <summary>临近天数：结束日期在此天数内视为临近</summary>
        public int WarningDaysBefore { get; set; } = 3;

        /// <summary>警告包含已逾期未完成的计划</summary>
        public bool WarningIncludeOverdue { get; set; } = true;

        /// <summary>同一对象同类邮件的重复发送间隔（天）</summary>
        public int DedupeDays { get; set; } = 1;

        /// <summary>各类型是否抄送管理员：键为类型代码</summary>
        public Dictionary<string, bool> CcAdmins { get; set; } = [];

        /// <summary>
        /// 该类型是否抄送所有管理员（可在设置中按类型管理）
        /// </summary>
        public bool ShouldCcAdmins(MailTypeDefinition type)
        {
            if (type == null || CcAdmins == null)
            {
                return false;
            }
            return CcAdmins.TryGetValue(type.Code, out bool value) && value;
        }

        public void SetCcAdmins(MailTypeDefinition type, bool value)
        {
            CcAdmins ??= [];
            CcAdmins[type.Code] = value;
        }

        /// <summary>邮件模板：键为 "&lt;类型代码&gt;.subject" / "&lt;类型代码&gt;.body"</summary>
        public Dictionary<string, string> Templates { get; set; } = [];

        /// <summary>
        /// 取模板内容（未自定义时返回 null，由发送时回退到内置默认模板）
        /// </summary>
        public string GetTemplate(MailTypeDefinition type, bool subject)
        {
            if (type == null || Templates == null)
            {
                return null;
            }
            string key = type.Code + (subject ? ".subject" : ".body");
            return Templates.TryGetValue(key, out string value) && !string.IsNullOrWhiteSpace(value) ? value : null;
        }

        public void SetTemplate(MailTypeDefinition type, bool subject, string value)
        {
            Templates ??= [];
            Templates[type.Code + (subject ? ".subject" : ".body")] = value;
        }

        public IReadOnlyList<string> SecurityOptions { get; } = ["None", "StartTls", "Ssl"];
    }

    /// <summary>
    /// UI 相关设置
    /// </summary>
    public class UiSettings
    {
        /// <summary>
        /// 全局字体族名称
        /// </summary>
        public string FontFamily { get; set; } = "Microsoft YaHei UI";

        /// <summary>
        /// 全局字体大小（像素）
        /// </summary>
        public double FontSize { get; set; } = 14;

        /// <summary>
        /// 全局字重（Normal/Medium/SemiBold/Bold），默认常规。
        /// 注意：部分中文字体（如微软雅黑 UI）没有 Medium 字面，选 Medium 会被系统合成为粗体。
        /// </summary>
        public string FontWeight { get; set; } = "Normal";

        /// <summary>
        /// Toast 提示出现位置（TopRight/TopLeft/BottomRight/BottomLeft），默认右上角
        /// </summary>
        public string ToastPosition { get; set; } = "TopRight";
    }

    /// <summary>
    /// 路径相关设置（保存到数据库的默认打开目录；ATE/EMI/数据库路径保存在程序目录本地文件）
    /// </summary>
    public class PathSettings
    {
        /// <summary>
        /// 计划表路径（导入计划表默认目录）
        /// </summary>
        public string SchedulePath { get; set; }

        /// <summary>
        /// 领用表路径（导入领用表默认目录）
        /// </summary>
        public string RequisitionPath { get; set; }

        /// <summary>
        /// 报告路径（一键报告默认打开目录）
        /// </summary>
        public string ReportPath { get; set; }
    }
}

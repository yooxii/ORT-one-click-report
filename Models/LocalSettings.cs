using System;
using System.Collections.Generic;

namespace ORT一键报告.Models
{
    /// <summary>
    /// 本机设置（程序目录 Data\local_settings.json，**只有一个文件**）：
    /// 只保存"必须留在本机"的设置项——
    /// 1) 数据文件夹位置（数据库/附件/配图等全部数据都放在那里，可在远程共享上）；
    /// 2) ATE / EMI 源数据路径（各台电脑自己的目录）；
    /// 3)（旧位置遗留、仅用于迁移）登录 cookie 字段——凭据已改存到当前用户目录 %LocalAppData%（见 LoginCredentialStore）；
    /// 4) 计划表列布局；
    /// 5) 最近搜索记忆（搜索框的历史关键字）。
    /// 其余设置（界面、邮件、业务路径等）随数据库放在数据文件夹里，多台电脑共用。
    /// 反序列化时缺失的项一律保持默认值，因此旧版本文件（或手工精简过的文件）都能读。
    /// </summary>
    public class LocalSettings
    {
        /// <summary>
        /// 数据文件夹：数据库、附件（OleFiles）、配图（PlanImages）等全部业务数据都在这里，
        /// 可以是本地目录，也可以是 \\服务器\共享\目录 之类的 UNC 路径；为空表示程序目录\Data。
        /// </summary>
        public string DataFolder { get; set; }

        /// <summary>ATE 源数据路径（本机目录）</summary>
        public string AteDataPath { get; set; }

        /// <summary>EMI 源数据路径（本机目录）</summary>
        public string EmiDataPath { get; set; }

        /// <summary>【旧位置遗留、仅用于迁移】登录用户名；凭据已改存用户目录，迁移后为 null</summary>
        public string LoginUsername { get; set; }

        /// <summary>【旧位置遗留、仅用于迁移】DPAPI 加密后的登录密码；凭据已改存用户目录，迁移后为 null</summary>
        public string LoginPasswordEnc { get; set; }

        /// <summary>【旧位置遗留、仅用于迁移】登录到期时间；凭据已改存用户目录，迁移后为 null</summary>
        public DateTime? LoginExpiry { get; set; }

        /// <summary>计划表列布局（"requisitions" / "plans" → 列键列表）</summary>
        public Dictionary<string, List<string>> PlansLayout { get; set; }

        /// <summary>最近搜索记忆（领退表/计划表共用搜索框，最新在前，最多 10 条）</summary>
        public List<string> PlanSearchHistory { get; set; }

        /// <summary>后台运行设置（本机：最小化到后台 / 开机自启）</summary>
        public BackgroundSettings Background { get; set; } = new BackgroundSettings();

        /// <summary>领退和计划窗口的页签布局（哪些页签显示、当前页签、浮在窗口外的页签与位置）</summary>
        public PlansTabsLayout PlansTabs { get; set; } = new PlansTabsLayout();

        /// <summary>预领的更新提醒设置（每周排程提醒 / 每天看板提醒，本机保存）</summary>
        public PreIssueReminderSettings PreIssueReminders { get; set; } = new PreIssueReminderSettings();

        /// <summary>MES（SIM44 回线工令）连接设置（本机保存：不同厂区/测试库可能用不同别名）</summary>
        public Sim44Settings Sim44 { get; set; } = new Sim44Settings();

        /// <summary>
        /// SP程序集里各程式 exe 的路径覆盖（条目 id → 完整路径）。
        /// 默认路径写在 <see cref="MesProgramInfo.DefaultPath"/>，厂区不一样时改这里，不动代码。
        /// </summary>
        public Dictionary<string, string> SpPrograms { get; set; }
    }

    /// <summary>
    /// MES 回线工令（SIM44 表）的连接设置，存本机：
    /// 数据库别名走 Oracle 的 tnsnames（默认与 BDE 别名 BSFC 指向同一个库），
    /// 表名允许带 schema 前缀（例如 APITECH.SIM44），桥接程序路径用于非常规安装位置。
    /// </summary>
    public class Sim44Settings
    {
        /// <summary>Oracle 连接别名（tnsnames.ora 里的名字）</summary>
        public string DataSource { get; set; } = "NT_BSFC.WORLD";

        /// <summary>表名（可带 schema，默认 SIM44）</summary>
        public string Table { get; set; } = "SIM44";

        /// <summary>桥接程序完整路径（留空则在程序目录下自动查找）</summary>
        public string BridgePath { get; set; }

        /// <summary>上次使用的数据库账号（只记账号，密码单独加密存用户目录）</summary>
        public string LastUser { get; set; }
    }

    /// <summary>
    /// 领退和计划窗口的外层页签布局（本机保存）：
    /// 三个页签（领退表 / 计划表 / 预领）可以显示或隐藏、可以拖动改顺序、可以浮在窗口外单独摆放，
    /// 关闭窗口时把状态记到这里，下次打开按记忆恢复。
    /// </summary>
    public class PlansTabsLayout
    {
        /// <summary>页签的排列顺序（拖动页签改的顺序；缺的页签按默认顺序补在后面）</summary>
        public List<string> Order { get; set; }

        /// <summary>当前显示的页签 id（requisitions / plans / preissue），按顺序排列</summary>
        public List<string> Visible { get; set; }

        /// <summary>当前选中的页签 id</summary>
        public string Active { get; set; }

        /// <summary>浮在窗口外的页签 id</summary>
        public List<string> Floating { get; set; }

        /// <summary>浮动窗口的位置与大小（页签 id → "x,y,width,height"）</summary>
        public Dictionary<string, string> Bounds { get; set; }
    }

    /// <summary>
    /// 预领的更新提醒设置（本机保存）：
    /// 月排程每周固定时间提醒导入更新，ORT 看板每天固定时间提醒导入更新。
    /// 时间用 HH:mm 文本保存，星期用 DayOfWeek 名称保存，避免跨版本序列化问题。
    /// </summary>
    public class PreIssueReminderSettings
    {
        /// <summary>每周排程提醒开关</summary>
        public bool WeeklyEnabled { get; set; } = true;

        /// <summary>每周排程提醒的星期（Monday..Sunday）</summary>
        public string WeeklyDay { get; set; } = "Monday";

        /// <summary>每周排程提醒的时间（HH:mm）</summary>
        public string WeeklyTime { get; set; } = "08:30";

        /// <summary>每天看板提醒开关</summary>
        public bool DailyEnabled { get; set; } = true;

        /// <summary>每天看板提醒的时间（HH:mm）</summary>
        public string DailyTime { get; set; } = "08:30";

        /// <summary>上一次排程提醒的日期（yyyy-MM-dd），同一天只提醒一次</summary>
        public string LastWeeklyPrompt { get; set; }

        /// <summary>上一次看板提醒的日期（yyyy-MM-dd），同一天只提醒一次</summary>
        public string LastDailyPrompt { get; set; }
    }

    /// <summary>
    /// 后台运行设置（保存在本机 local_settings.json）：
    /// 「最小化到后台」只影响本机的窗口关闭行为，「开机自启」要写当前用户的注册表启动项，
    /// 因此这两项都跟随本机，不随数据库共享给其他电脑。
    /// 远程路径（网络共享 / 网络盘 / SUBST 虚拟盘）里的程序不登记开机自启，见 StartupManager。
    /// </summary>
    public class BackgroundSettings
    {
        /// <summary>关闭主窗口时最小化到后台（托盘）；未勾选则直接关闭程序。默认勾选</summary>
        public bool MinimizeToTrayOnClose { get; set; } = true;

        /// <summary>开机自启（当前用户，无需管理员权限）</summary>
        public bool AutoStart { get; set; }

        /// <summary>开机自启时直接进后台（托盘），不弹出主窗口</summary>
        public bool AutoStartToBackground { get; set; }
    }
}

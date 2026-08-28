using FreeSql.DataAnnotations;

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
        /// 路径设置
        /// </summary>
        public PathSettings Paths { get; set; } = new PathSettings();
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

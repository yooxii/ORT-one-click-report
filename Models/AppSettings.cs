namespace ORT一键报告.Models
{
    /// <summary>
    /// 应用设置（以 JSON 格式保存到 Data\settings.json）
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
    /// 路径相关设置（各对话框默认打开目录）
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

        /// <summary>
        /// ATE数据路径（ATE数据默认打开目录）
        /// </summary>
        public string AteDataPath { get; set; }

        /// <summary>
        /// EMI数据路径（EMI数据默认打开目录）
        /// </summary>
        public string EmiDataPath { get; set; }
    }
}

namespace ORT一键报告.Models
{
    /// <summary>
    /// 报告状态（Plan.ReportStatus）的取值：
    /// - <see cref="Complete"/> / <see cref="InProgress"/> 由扫描报告文件夹时按 TestStatus 表自动写入；
    /// - <see cref="NotRequired"/> 只能由用户在编辑窗口手工设置，设置后扫描不再覆盖该记录。
    /// 值直接落库为字符串，界面显示与存储一致（无需再做映射）。
    /// </summary>
    public static class ReportStatusKind
    {
        /// <summary>已完成：TestStatus 表里每个测试项都满足 UNITS UNDER TEST == TOTAL + FAIL</summary>
        public const string Complete = "已完成";

        /// <summary>进行中：TestStatus 表里存在测试项不满足完成条件（含全部为 0 的未开始情况）</summary>
        public const string InProgress = "进行中";

        /// <summary>无要求：用户手工覆盖，扫描不再自动改动</summary>
        public const string NotRequired = "无要求";

        /// <summary>下拉框选项（按"自动可产出 → 用户专属"排序）</summary>
        public static readonly string[] All = [Complete, InProgress, NotRequired];

        /// <summary>
        /// 该值是否被用户锁定（扫描不再覆盖）：仅「无要求」锁定；
        /// null / 空串 / 已完成 / 进行中 都允许扫描按 TestStatus 结果重写。
        /// </summary>
        public static bool IsUserLocked(string value) => value == NotRequired;

        /// <summary>
        /// 判断值是否合法（null/空串视为"尚未设置"，也合法）
        /// </summary>
        public static bool IsValid(string value)
            => string.IsNullOrWhiteSpace(value)
               || value == Complete
               || value == InProgress
               || value == NotRequired;
    }
}

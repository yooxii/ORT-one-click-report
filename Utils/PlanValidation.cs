using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ORT一键报告.Utils
{
    /// <summary>
    /// 计划状况归类：把库里的状况文本（Ongoing / Close / Pending，历史上也有 Closed 等写法）
    /// 归成三类，供计划表状况列配色与搜索栏状况统计共用。
    /// </summary>
    public static class PlanStatusKind
    {
        /// <summary>测试中</summary>
        public const string Ongoing = "ongoing";

        /// <summary>预排测试（未开始）</summary>
        public const string Pending = "pending";

        /// <summary>已结案</summary>
        public const string Closed = "closed";

        /// <summary>
        /// 归类状况文本；认不出（空值或自定义写法）返回空字符串，调用方按"不配色/不统计"处理。
        /// 判定不区分大小写，且 Close / Closed / 结案 都算已结案。
        /// </summary>
        public static string Of(string status)
        {
            if (string.IsNullOrWhiteSpace(status))
            {
                return "";
            }
            string text = status.Trim().ToLowerInvariant();
            if (text.Contains("ongoing") || text.Contains("进行中") || text.Contains("進行中") || text.Contains("測試中") || text.Contains("测试中"))
            {
                return Ongoing;
            }
            if (text.Contains("pending") || text.Contains("待测") || text.Contains("待測") || text.Contains("預排") || text.Contains("预排"))
            {
                return Pending;
            }
            if (text.Contains("close") || text.Contains("结案") || text.Contains("結案") || text.Contains("已完成"))
            {
                return Closed;
            }
            return "";
        }
    }

    /// <summary>
    /// 计划数据编辑校验：指定列的格式限制与字典约束
    /// </summary>
    public static class PlanValidation
    {
        /// <summary>
        /// 状况只允许三种值
        /// </summary>
        public static readonly string[] ValidStatuses = ["Ongoing", "Close", "Pending"];

        /// <summary>
        /// 工作编号格式：以 QRT 或 RT 开头，之后跟 4 位年月（如 2608），最后是至少两位、从 01 开始的编号
        /// </summary>
        private static readonly Regex JobNoRegex = new(@"^(QRT|RT)(\d{4})(\d{2,})$", RegexOptions.IgnoreCase);

        /// <summary>
        /// 校验工作编号；合法返回null，否则返回错误描述
        /// </summary>
        public static string ValidateJobNo(string jobNo)
        {
            if (string.IsNullOrWhiteSpace(jobNo))
            {
                return null; // 允许为空（由必填校验另行处理）
            }
            Match m = JobNoRegex.Match(jobNo.Trim());
            if (!m.Success)
            {
                return "工作编号格式应为：QRT/RT + 4位年月 + 至少2位编号（如 RT260801）";
            }
            int seq = int.Parse(m.Groups[3].Value);
            if (seq < 1)
            {
                return "工作编号末尾编号必须从 01 开始";
            }
            return null;
        }

        /// <summary>
        /// 校验状况值；合法返回null
        /// </summary>
        public static string ValidateStatus(string status)
        {
            if (string.IsNullOrWhiteSpace(status))
            {
                return null;
            }
            return ValidStatuses.Any(s => s.Equals(status.Trim(), System.StringComparison.OrdinalIgnoreCase))
                ? null
                : $"状况只能是：{string.Join(" / ", ValidStatuses)}";
        }

        /// <summary>
        /// 校验值必须在字典中（测试项目/产品别/客户别/阶段）；空值合法，返回null
        /// </summary>
        public static string ValidateInCatalog(string value, IEnumerable<string> catalog, string columnName)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }
            string trimmed = value.Trim();
            return catalog.Any(c => c == trimmed)
                ? null
                : $"{columnName} [{trimmed}] 不在字典中，请先在\"管理\"模块中添加";
        }
    }
}

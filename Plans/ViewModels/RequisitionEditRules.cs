using ORT一键报告.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ORT一键报告.Plans.ViewModels
{
    /// <summary>
    /// 领退表新增/编辑窗口（WindowRequisitionEdit）的纯规则：
    /// 单体去向按机种名称判定、完工令（Work Order）里解析线别与 D/C、以及「保存并继续」时编号的递增避重。
    /// 不碰数据库与界面，便于单元测试。
    /// </summary>
    public static class RequisitionEditRules
    {
        /* ###############################  单体去向  ################################ */

        /// <summary>
        /// 单体去向默认判定：机种名称以 W 开头（不区分大小写，忽略首尾空白）→ 报废，其余（含空值）→ 入库。
        /// </summary>
        public static string DispositionFromModel(string modelName)
            => (modelName ?? "").Trim().StartsWith("W", StringComparison.OrdinalIgnoreCase)
                ? RequisitionDispositionKind.Scrap
                : RequisitionDispositionKind.StockIn;

        /* ###############################  完工令解析  ################################ */

        /// <summary>
        /// 线别：完工令「倒数第六位起的三位」；不足 6 位返回 null（取不到就不补全）。
        /// </summary>
        public static string LineNoFromWorkOrder(string workOrder)
        {
            string wo = (workOrder ?? "").Trim();
            return wo.Length >= 6 ? wo.Substring(wo.Length - 6, 3) : null;
        }

        /// <summary>
        /// D/C（界面上显示为「周期」）：年份后两位 + 完工令「倒数第三位起的两位」，
        /// 例如 26 年第 33 周 → 2633；完工令不足 3 位返回 null。
        /// </summary>
        public static string DcFromWorkOrder(string workOrder, int year)
        {
            string wo = (workOrder ?? "").Trim();
            return wo.Length >= 3
                ? (year % 100).ToString("D2") + wo.Substring(wo.Length - 3, 2)
                : null;
        }

        /* ###############################  机种名称补全  ################################ */

        /// <summary>
        /// 机种名称补全候选：已有名称里第一个「以已输入内容开头且比它更长」的（不区分大小写、忽略首尾空白）。
        /// 命中多个时取最短的（最贴近已输入内容），长度相同按字典序，保证结果稳定；没有候选返回 null。
        /// </summary>
        public static string ModelSuggestion(IEnumerable<string> candidates, string typed)
        {
            string text = (typed ?? "").Trim();
            if (candidates == null || text.Length == 0)
            {
                return null;
            }
            return candidates.Where(c => !string.IsNullOrWhiteSpace(c))
                .Select(c => c.Trim())
                .Where(c => c.Length > text.Length && c.StartsWith(text, StringComparison.OrdinalIgnoreCase))
                .OrderBy(c => c.Length)
                .ThenBy(c => c, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault();
        }

        /* ###############################  编号递增  ################################ */

        private static readonly Regex JobNoPattern = new(@"^(?<head>(?:QRT|RT)\d{4})(?<seq>\d+)$", RegexOptions.IgnoreCase);

        private static readonly Regex ReturnRtPattern = new(@"^(?<head>RTAH\d{4})(?<seq>\d+)$", RegexOptions.IgnoreCase);

        /// <summary>
        /// 工作编号（QRT/RT + 年月 + 序号）序号 +1；格式认不出返回 null。
        /// </summary>
        public static string NextJobNo(string jobNo) => NextInSequence(jobNo, JobNoPattern);

        /// <summary>
        /// 回线RT工令（RTAH + 年月 + 序号）序号 +1；格式认不出返回 null。
        /// </summary>
        public static string NextReturnRtOrder(string returnRtOrder) => NextInSequence(returnRtOrder, ReturnRtPattern);

        /// <summary>
        /// 按「前缀 + 末尾序号」把序号加一：序号 ≥ 100 时按实际位数展开（与自动编号的一致约定），否则补到两位。
        /// </summary>
        private static string NextInSequence(string value, Regex pattern)
        {
            Match match = pattern.Match((value ?? "").Trim());
            if (!match.Success || !int.TryParse(match.Groups["seq"].Value, out int sequence))
            {
                return null;
            }
            int next = sequence + 1;
            return match.Groups["head"].Value + (next >= 100 ? next.ToString() : next.ToString("D2"));
        }
    }
}

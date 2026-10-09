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

        /* ###############################  编号体检  ################################ */

        /// <summary>
        /// 取出回线RT工令里的年月段（RTAH 之后的四位 yyMM）；不是「RTAH + 四位年月 + 序号」的形状返回 null。
        /// </summary>
        public static string ReturnRtYearMonth(string returnRtOrder)
        {
            Match match = ReturnRtPattern.Match((returnRtOrder ?? "").Trim());
            return match.Success ? match.Groups["head"].Value.Substring("RTAH".Length) : null;
        }

        /// <summary>
        /// 体检单条回线RT工令，返回（问题类型, 建议编号）；正常返回 null：
        /// - 空值 → 正常（还没登记回线，编号本来就可以为空）；
        /// - 不是「RTAH + 四位年月 + 序号」的形状 → 格式无法识别（建议编号为 null）；
        /// - 年月段与领用日期、回线日期**都对不上** → 月段不符（建议编号保留原序号、换成领用日期的年月）；
        /// - 两个日期都为空时无从判断年月，只做格式检查。
        ///
        /// 之所以「对上一个就算正常」：历史数据里两种口径都存在——
        /// 有的按领用日期编号（与本程序现在的自动编号一致），有的按回线日期编号（3 月领用、4 月回线就是 4 月的号）。
        /// 只比领用日期会把后一种全判成异常；实测 186 条里「与回线日期一致」89 条、「与领用日期一致」82 条，
        /// 两个都对不上的才是真异常。
        /// </summary>
        public static (ReturnRtCodeIssueKind Kind, string SuggestedCode)? InspectReturnRt(
            string returnRtOrder, DateTime? requisitionDate, DateTime? returnDate)
        {
            string value = (returnRtOrder ?? "").Trim();
            if (value.Length == 0)
            {
                return null;
            }
            Match match = ReturnRtPattern.Match(value);
            if (!match.Success)
            {
                return (ReturnRtCodeIssueKind.UnrecognizedFormat, null);
            }
            string codeYearMonth = match.Groups["head"].Value.Substring("RTAH".Length);
            string requisitionYearMonth = requisitionDate?.ToString("yyMM");
            string returnYearMonth = returnDate?.ToString("yyMM");
            if (requisitionYearMonth == null && returnYearMonth == null)
            {
                return null;
            }
            if (string.Equals(codeYearMonth, requisitionYearMonth, StringComparison.Ordinal)
                || string.Equals(codeYearMonth, returnYearMonth, StringComparison.Ordinal))
            {
                return null;
            }
            // 建议编号按领用日期（与本程序自动编号的口径一致）；没有领用日期时才退回回线日期
            string useYearMonth = requisitionYearMonth ?? returnYearMonth;
            // 只换年月段、原样保留序号位数（001 这种三位写法也照旧）
            return (ReturnRtCodeIssueKind.MonthMismatch, "RTAH" + useYearMonth + match.Groups["seq"].Value);
        }
    }

    /// <summary>回线RT工令编号体检发现的问题类型</summary>
    public enum ReturnRtCodeIssueKind
    {
        /// <summary>编号里的年月段与领用日期、回线日期都对不上（如 1 月的记录写成 RTAH2610…）</summary>
        MonthMismatch,

        /// <summary>值不是「RTAH + 四位年月 + 序号」的形状（例如填成了说明文字）</summary>
        UnrecognizedFormat
    }

    /// <summary>一条回线RT工令编号体检结果（只读展示用）</summary>
    public sealed class ReturnRtCodeIssue
    {
        public long Id { get; set; }
        public DateTime? RequisitionDate { get; set; }

        /// <summary>回线日期（历史数据里有的按它编号，判异常时要一起看）</summary>
        public DateTime? ReturnDate { get; set; }

        public string RequisitionNo { get; set; }
        public string ModelName { get; set; }
        public string ReturnRtOrder { get; set; }
        public ReturnRtCodeIssueKind Kind { get; set; }

        /// <summary>编号里实际写的年月（yyMM；格式认不出为 null）</summary>
        public string CodeYearMonth { get; set; }

        /// <summary>建议改成这个编号（仅「月段不符」时有值）</summary>
        public string SuggestedCode { get; set; }
    }
}

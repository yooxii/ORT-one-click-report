using System;
using System.Collections.Generic;
using System.Linq;

namespace ORT一键报告.Models
{
    /// <summary>
    /// 测试种类（测试项目的归类）：报告 ORT Plan 的分类行与 TestStatus 的分组行都按它来分。
    /// 计划索引会按历史报告里该测试项所在的分类自动归类；认不出来的归到"不确定"，
    /// 由用户在"管理 → 测试项目"里手工归类。
    /// </summary>
    public static class TestCategories
    {
        /// <summary>可靠性与环境类（历史报告里记为 RELIABILITY TEST，TestStatus 里显示 ENVIRONMENT TESTS）</summary>
        public const string Reliability = "RELIABILITY TEST";

        /// <summary>电磁兼容类</summary>
        public const string Emc = "EMC";

        /// <summary>认不出所属种类时的归类</summary>
        public const string Uncertain = "不确定";

        /// <summary>界面下拉可选的测试种类（顺序即下拉顺序）</summary>
        public static readonly string[] Known = [Reliability, Emc, Uncertain];

        /// <summary>EMC 关键词（先判 EMC：Conducted EMI Measurement 之类不能被"环境"关键词抢走）</summary>
        private static readonly string[] EmcKeywords =
        [
            "EMI", "EMC", "ESD", "SURGE", "EFT", "RADIATED", "CONDUCTED", "HARMONIC",
            "FLICKER", "LISN", "61000", "POWER LINE", "ANTENNA", "ELECTROSTATIC",
            "FAST TRANSIENT", "BURST", "IMMUNITY"
        ];

        /// <summary>可靠性/环境类关键词</summary>
        private static readonly string[] ReliabilityKeywords =
        [
            "THERMAL", "SHOCK", "BURN", "VIBRATION", "DROP", "HUMID", "TEMPERATURE", "TEMP",
            "ALTITUDE", "DUST", "SALT", "SPRAY", "CYCLE", "CYCLING", "AGING", "AGEING",
            "LIFE", "ENDURANCE", "COLD", "HEAT", "ACOUSTIC", "NOISE", "MECHANICAL",
            "PACKAGE", "IMPACT", "STRESS", "FALL", "WATER", "IPX", "HALT", "MTBF",
            "TUMBLE", "STRAIN", "ABRASION", "RELIABILITY", "REL TEST"
        ];

        /// <summary>
        /// 归类：优先按关键词判断（EMC → 可靠性 → 不确定）
        /// </summary>
        public static string Classify(string testItemName)
        {
            string text = (testItemName ?? "").ToUpperInvariant();
            if (text.Length == 0)
            {
                return Uncertain;
            }
            if (EmcKeywords.Any(text.Contains))
            {
                return Emc;
            }
            if (ReliabilityKeywords.Any(text.Contains))
            {
                return Reliability;
            }
            return Uncertain;
        }

        /// <summary>
        /// 把历史报告/计划里的分类文本（RELIABILITY TEST / ENVIRONMENT TESTS / EMC 等）归一成本程序的测试种类；
        /// 认不出来返回 null
        /// </summary>
        public static string Normalize(string raw)
        {
            string text = (raw ?? "").Trim().ToUpperInvariant();
            if (text.Length == 0)
            {
                return null;
            }
            if (text.Contains("EMC") || text.Contains("EMI"))
            {
                return Emc;
            }
            if (text.Contains("RELIABILITY") || text.Contains("ENVIRONMENT") || text.Contains("ENVIRONMENTAL"))
            {
                return Reliability;
            }
            if (text.Contains("MECHANICAL"))
            {
                return Reliability;
            }
            if (text.Contains("SAFETY"))
            {
                return Reliability;
            }
            return null;
        }

        /// <summary>
        /// 从一组历史分类里取出现次数最多且能识别的那个（用于计划索引归并）
        /// </summary>
        public static string Normalize(IEnumerable<string> raws)
        {
            return (raws ?? [])
                .Select(Normalize)
                .Where(c => !string.IsNullOrEmpty(c))
                .GroupBy(c => c)
                .OrderByDescending(g => g.Count())
                .Select(g => g.Key)
                .FirstOrDefault();
        }

        /// <summary>显示名（TestStatus 分组行的写法与历史报告一致：ENVIRONMENT TESTS / EMC ）</summary>
        public static string StatusDisplay(string category)
        {
            string normalized = Normalize(category);
            if (normalized == Reliability)
            {
                return "ENVIRONMENT TESTS";
            }
            if (normalized == Emc)
            {
                return "EMC ";
            }
            return string.IsNullOrWhiteSpace(category) ? Uncertain : category.Trim();
        }
    }
}

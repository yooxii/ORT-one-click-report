using System;
using System.Collections.Generic;
using System.Linq;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 预领·同系列客户优先级规则：
    /// 机种形状「前 6 位-后 4 位」（如 FSL020-00AG）= 系列（前 6 位）+ 客户代码（后段前 2 位）+ 版本（后段后 2 位），
    /// **同一个系列只需要开一张单**，挑客户代码优先级最高的那个机种去开即可。
    /// 优先级用**数字**表示：数字越小越优先，可以填相同的数字表示并列
    /// （并列时再按客户代码、机种名排序，保证结果稳定）；没列出的客户代码、以及拆不出客户代码的机种都排最后。
    /// </summary>
    public sealed class PreIssueSeriesPriority
    {
        /// <summary>默认优先级：GA / 4C / 93 优先，00（标准）最后</summary>
        private static readonly CustomerPriorityEntry[] Defaults =
        [
            new() { Code = "GA", Priority = 1 },
            new() { Code = "4C", Priority = 2 },
            new() { Code = "93", Priority = 3 },
            new() { Code = "00", Priority = 9 }
        ];

        private readonly List<CustomerPriorityEntry> _entries;
        private readonly Dictionary<string, int> _priorities = new(StringComparer.OrdinalIgnoreCase);

        public PreIssueSeriesPriority(IEnumerable<CustomerPriorityEntry> entries)
        {
            List<CustomerPriorityEntry> normalized = Normalize(entries);
            if (normalized.Count == 0)
            {
                normalized = Normalize(Defaults);
            }
            _entries = normalized
                .OrderBy(e => e.Priority)
                .ThenBy(e => e.Code, StringComparer.OrdinalIgnoreCase)
                .ToList();
            foreach (CustomerPriorityEntry entry in _entries)
            {
                if (!_priorities.ContainsKey(entry.Code))
                {
                    _priorities[entry.Code] = entry.Priority;
                }
            }
        }

        /// <summary>当前生效的优先级（已按数字从优到劣排好）</summary>
        public IReadOnlyList<CustomerPriorityEntry> Entries => _entries;

        /// <summary>默认优先级</summary>
        public static PreIssueSeriesPriority CreateDefault() => new(Defaults);

        /// <summary>默认优先级条目（每次返回副本，避免被调用方改动）</summary>
        public static List<CustomerPriorityEntry> DefaultEntries() =>
            Defaults.Select(e => new CustomerPriorityEntry { Code = e.Code, Priority = e.Priority }).ToList();

        /// <summary>归一化条目：去空白、代码转大写、同一代码只留第一条</summary>
        public static List<CustomerPriorityEntry> Normalize(IEnumerable<CustomerPriorityEntry> entries)
        {
            List<CustomerPriorityEntry> result = [];
            if (entries == null)
            {
                return result;
            }
            foreach (CustomerPriorityEntry raw in entries)
            {
                string code = (raw?.Code ?? "").Trim().ToUpperInvariant();
                if (code.Length == 0 || result.Any(e => string.Equals(e.Code, code, StringComparison.Ordinal)))
                {
                    continue;
                }
                result.Add(new CustomerPriorityEntry { Code = code, Priority = raw.Priority });
            }
            return result;
        }

        /// <summary>
        /// 解析库里存的文本：每项写成「代码:数字」（也接受 = 分隔、允许等号两边有空格），
        /// 兼容旧格式「GA,4C,93,00」（没写数字时按出现顺序 1、2、3… 当优先级）。
        /// </summary>
        public static List<CustomerPriorityEntry> ParseStored(string text)
        {
            List<CustomerPriorityEntry> entries = [];
            if (string.IsNullOrWhiteSpace(text))
            {
                return entries;
            }
            // 「GA = 3」这种写法先把等号/冒号两边的空格去掉，再按逗号/分号/竖线/空白拆
            string normalized = System.Text.RegularExpressions.Regex.Replace(text.Trim(), @"\s*([:=])\s*", "$1");
            string[] tokens = normalized.Split([',', ';', '|', ' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
            for (int index = 0; index < tokens.Length; index++)
            {
                string token = tokens[index].Trim();
                int separator = token.IndexOfAny([':', '=']);
                string code = separator > 0 ? token.Substring(0, separator).Trim() : token;
                if (code.Length == 0)
                {
                    continue;
                }
                int priority = index + 1;
                if (separator > 0 && int.TryParse(token.Substring(separator + 1).Trim(), out int parsed))
                {
                    priority = parsed;
                }
                entries.Add(new CustomerPriorityEntry { Code = code.ToUpperInvariant(), Priority = priority });
            }
            return Normalize(entries);
        }

        /// <summary>序列化成库里存的文本（按优先级排好，形如 GA:1,4C:2）</summary>
        public static string Format(IEnumerable<CustomerPriorityEntry> entries) =>
            string.Join(",", Normalize(entries)
                .OrderBy(e => e.Priority)
                .ThenBy(e => e.Code, StringComparer.OrdinalIgnoreCase)
                .Select(e => e.Code + ":" + e.Priority));

        /// <summary>可读文本（GA=1、4C=1、93=2、00=9），界面提示用</summary>
        public static string Describe(IEnumerable<CustomerPriorityEntry> entries)
        {
            List<CustomerPriorityEntry> list = Normalize(entries)
                .OrderBy(e => e.Priority)
                .ThenBy(e => e.Code, StringComparer.OrdinalIgnoreCase)
                .ToList();
            return list.Count == 0 ? "（空）" : string.Join("、", list.Select(e => $"{e.Code}={e.Priority}"));
        }

        /// <summary>
        /// 系列键：能拆出「前 6 位-后段」的取前 6 位（忽略大小写）；
        /// 形状不对的机种用整串，等于「只有同名机种才算同系列」。
        /// </summary>
        public static string SeriesKeyOf(string model)
        {
            string text = (model ?? "").Trim();
            if (text.Length == 0)
            {
                return "";
            }
            return OrtBlacklistMatcher.TryParseSeries(text, out string series, out _, out _)
                ? series.ToUpperInvariant()
                : text.ToUpperInvariant();
        }

        /// <summary>客户代码（后 4 位的前 2 位）；机种拆不出来、或后段不是 4 位时返回 null</summary>
        public static string CustomerCodeOf(string model) =>
            OrtBlacklistMatcher.TryParseSeries((model ?? "").Trim(), out _, out string customer, out _)
                ? customer?.ToUpperInvariant()
                : null;

        /// <summary>系列键（实例版，供服务里按系列归堆用，等价于 <see cref="SeriesKeyOf"/>)</summary>
        public string SeriesKey(string model) => SeriesKeyOf(model);

        /// <summary>机种的优先级数字（越小越优先；没列出的客户代码、拆不出代码的机种排最后）</summary>
        public int RankOf(string model)
        {
            string code = CustomerCodeOf(model);
            return code != null && _priorities.TryGetValue(code, out int priority) ? priority : int.MaxValue;
        }

        /// <summary>同一个系列里挑要开单的那个机种（优先级数字最小；并列时按客户代码、机种名排序）</summary>
        public string PickRepresentative(IEnumerable<string> models) => (models ?? [])
            .Where(m => !string.IsNullOrWhiteSpace(m))
            .Select(m => new { Model = m.Trim(), Rank = RankOf(m), Code = CustomerCodeOf(m) ?? "" })
            .OrderBy(x => x.Rank)
            .ThenBy(x => x.Code, StringComparer.OrdinalIgnoreCase)
            .ThenBy(x => x.Model, StringComparer.OrdinalIgnoreCase)
            .Select(x => x.Model)
            .FirstOrDefault();
    }

    /// <summary>客户代码 + 优先级数字（数字越小越优先，可以并列）</summary>
    public class CustomerPriorityEntry
    {
        /// <summary>客户代码（机种后段前 2 位）</summary>
        public string Code { get; set; }

        /// <summary>优先级数字：越小越优先，相同即并列</summary>
        public int Priority { get; set; }
    }
}

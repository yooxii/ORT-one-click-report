using System;
using System.Collections.Generic;
using System.Linq;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 预领·同系列客户优先级规则：
    /// 机种形状「前 6 位-后 4 位」（如 FSL020-00AG）= 系列（前 6 位）+ 客户代码（后段前 2 位）+ 版本（后段后 2 位），
    /// **同一个系列只需要开一张单**，挑客户代码优先级最高的那个机种去开即可。
    /// 优先级按配置里的顺序从高到低（默认 GA → 4C → 93 → 00，其中 00 = 标准、排在最后）；
    /// 没列出的客户代码、以及拆不出客户代码的机种都排最后，同档按机种名排序，保证结果稳定。
    /// </summary>
    public sealed class PreIssueSeriesPriority
    {
        /// <summary>默认优先级顺序：GA / 4C / 93 优先，00（标准）最后</summary>
        public static readonly string[] DefaultCodes = ["GA", "4C", "93", "00"];

        private readonly List<string> _codes;
        private readonly Dictionary<string, int> _ranks = new(StringComparer.OrdinalIgnoreCase);

        public PreIssueSeriesPriority(IEnumerable<string> codes)
        {
            _codes = Normalize(codes);
            if (_codes.Count == 0)
            {
                _codes = Normalize(DefaultCodes);
            }
            for (int index = 0; index < _codes.Count; index++)
            {
                if (!_ranks.ContainsKey(_codes[index]))
                {
                    _ranks[_codes[index]] = index;
                }
            }
        }

        /// <summary>当前生效的优先级（从高到低）</summary>
        public IReadOnlyList<string> Codes => _codes;

        /// <summary>默认优先级</summary>
        public static PreIssueSeriesPriority CreateDefault() => new(DefaultCodes);

        /// <summary>归一化代码列表：去空白、转大写、去重、保持顺序</summary>
        public static List<string> Normalize(IEnumerable<string> codes)
        {
            List<string> result = [];
            if (codes == null)
            {
                return result;
            }
            foreach (string raw in codes)
            {
                string code = (raw ?? "").Trim().ToUpperInvariant();
                if (code.Length == 0 || result.Contains(code, StringComparer.Ordinal))
                {
                    continue;
                }
                result.Add(code);
            }
            return result;
        }

        /// <summary>解析库里存的文本（逗号 / 分号 / 竖线 / 空白分隔）</summary>
        public static List<string> ParseStored(string text) =>
            Normalize(string.IsNullOrWhiteSpace(text)
                ? null
                : text.Split([',', ';', '|', ' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries));

        /// <summary>序列化成库里存的文本</summary>
        public static string Format(IEnumerable<string> codes) => string.Join(",", Normalize(codes));

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

        /// <summary>机种的优先级（越小越优先；没列出的客户代码、拆不出代码的机种排最后）</summary>
        public int RankOf(string model)
        {
            string code = CustomerCodeOf(model);
            return code != null && _ranks.TryGetValue(code, out int rank) ? rank : int.MaxValue;
        }

        /// <summary>同一个系列里挑要开单的那个机种（优先级最高；同档按客户代码、机种名排序）</summary>
        public string PickRepresentative(IEnumerable<string> models) => (models ?? [])
            .Where(m => !string.IsNullOrWhiteSpace(m))
            .Select(m => new { Model = m.Trim(), Rank = RankOf(m), Code = CustomerCodeOf(m) ?? "" })
            .OrderBy(x => x.Rank)
            .ThenBy(x => x.Code, StringComparer.OrdinalIgnoreCase)
            .ThenBy(x => x.Model, StringComparer.OrdinalIgnoreCase)
            .Select(x => x.Model)
            .FirstOrDefault();
    }
}

using NLog;
using ORT一键报告.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ORT一键报告.Services
{
    /// <summary>
    /// ORT 黑名单的匹配器：黑名单条目有两种——
    /// 1) 機種（精确）：整串忽略大小写相等；
    /// 2) 正则：按 .NET 正则做「包含匹配」（要整串匹配请自己写 ^…$），同样忽略大小写。
    /// 一个机种命中任意一条就算在黑名单里；管理窗口里写错的正则会在加载时被跳过（不会让整份黑名单失效）。
    /// </summary>
    public class OrtBlacklistMatcher
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly Dictionary<string, OrtBlacklistRule> _exact = new(StringComparer.OrdinalIgnoreCase);
        private readonly List<OrtBlacklistRule> _patterns = [];

        /// <summary>精确（機種）条目数</summary>
        public int ExactCount => _exact.Count;

        /// <summary>正则条目数</summary>
        public int PatternCount => _patterns.Count;

        /// <summary>总条目数</summary>
        public int Count => _exact.Count + _patterns.Count;

        /// <summary>从黑名单表构造（写错的正则跳过并记日志）</summary>
        public static OrtBlacklistMatcher Create(IEnumerable<OrtBlacklistItem> items)
        {
            OrtBlacklistMatcher matcher = new();
            foreach (OrtBlacklistItem item in items ?? [])
            {
                if (string.IsNullOrWhiteSpace(item.Model))
                {
                    continue;
                }
                matcher.Add(item.Model, item.IsPattern);
            }
            return matcher;
        }

        /// <summary>加一条规则（正则写错时返回 false，不加入）</summary>
        public bool Add(string value, bool isPattern)
        {
            string text = (value ?? "").Trim();
            if (text.Length == 0)
            {
                return false;
            }
            if (!isPattern)
            {
                _exact[text] = new OrtBlacklistRule { Value = text, IsPattern = false };
                return true;
            }
            if (!TryCompile(text, out Regex regex))
            {
                return false;
            }
            _patterns.Add(new OrtBlacklistRule { Value = text, IsPattern = true, Regex = regex });
            return true;
        }

        /// <summary>这个机种在不在黑名单里</summary>
        public bool Matches(string model) => Match(model) != null;

        /// <summary>命中黑名单的规则（界面提示"被哪条规则挡住"用），没有命中返回 null</summary>
        public OrtBlacklistRule Match(string model)
        {
            string text = (model ?? "").Trim();
            if (text.Length == 0)
            {
                return null;
            }
            if (_exact.TryGetValue(text, out OrtBlacklistRule exact))
            {
                return exact;
            }
            foreach (OrtBlacklistRule rule in _patterns)
            {
                try
                {
                    if (rule.Regex != null && rule.Regex.IsMatch(text))
                    {
                        return rule;
                    }
                }
                catch (RegexMatchTimeoutException)
                {
                    _logger.Warn($"黑名单正则匹配超时（规则：{rule.Value}，机种：{text}），本条按不命中处理");
                }
            }
            return null;
        }

        /// <summary>校验正则写法；写错时给出 .NET 的原始错误说明</summary>
        public static bool TryCompile(string pattern, out Regex regex, out string error)
        {
            regex = null;
            error = null;
            if (string.IsNullOrWhiteSpace(pattern))
            {
                error = "正则不能为空";
                return false;
            }
            try
            {
                regex = new Regex(pattern.Trim(), RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(500));
                return true;
            }
            catch (ArgumentException ex)
            {
                error = ex.Message;
                return false;
            }
        }

        private bool TryCompile(string pattern, out Regex regex) => TryCompile(pattern, out regex, out _);

        /* ------------------ 「机种 → 同系列正则」 ------------------ */

        /// <summary>
        /// 机种形状：前 6 位 + '-' + 一段字母数字（例：FSL020-00AG；现场也有 5 位的，如 WAS003-F90GB）。
        /// **前 6 位相同就算同一个系列**；后段正好 4 位时再拆出客户短码（前 2）与版本（后 2）。
        /// </summary>
        private static readonly Regex SeriesModelPattern =
            new(@"^([A-Za-z0-9]{6})-([A-Za-z0-9]{1,8})$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        /// <summary>机种是不是「前 6 位-后 4 位」的形状</summary>
        public static bool IsSeriesModel(string model) => TryParseSeries(model, out _, out _, out _);

        /// <summary>拆机种：系列（前 6 位）、客户短码（后 4 位前 2）、版本（后 4 位后 2）</summary>
        public static bool TryParseSeries(string model, out string series, out string customer, out string version)
        {
            series = null;
            customer = null;
            version = null;
            Match match = SeriesModelPattern.Match((model ?? "").Trim());
            if (!match.Success)
            {
                return false;
            }
            string prefix = match.Groups[1].Value;
            string suffix = match.Groups[2].Value;
            series = prefix;
            // 后段正好 4 位才按「客户短码 2 + 版本 2」拆；其它长度只给出系列
            if (suffix.Length == 4)
            {
                customer = suffix.Substring(0, 2);
                version = suffix.Substring(2, 2);
            }
            return true;
        }

        /// <summary>
        /// 把单个机种转成「同系列」正则：前 6 位相同、横线后面任意字母数字。
        /// 例：FSL020-00AG → ^FSL020-[A-Za-z0-9]+$（同系列里后段位数不一样的机种也能一起挡掉）
        /// </summary>
        public static bool TryBuildSeriesPattern(string model, out string pattern, out string series,
            out string customer, out string version)
        {
            pattern = null;
            if (!TryParseSeries(model, out series, out customer, out version))
            {
                return false;
            }
            pattern = "^" + Regex.Escape(series) + "-[A-Za-z0-9]+$";
            return true;
        }

        /// <summary>命中的规则说明（例如「正则 FSL0.*」）</summary>
        public static string Describe(OrtBlacklistRule rule) =>
            rule == null ? null : (rule.IsPattern ? "正则 " + rule.Value : rule.Value);
    }

    /// <summary>黑名单里的一条规则</summary>
    public class OrtBlacklistRule
    {
        /// <summary>機種或正则文本</summary>
        public string Value { get; set; }

        /// <summary>是否正则模式</summary>
        public bool IsPattern { get; set; }

        /// <summary>正则模式下的编译结果（精确模式为空）</summary>
        public Regex Regex { get; set; }
    }
}

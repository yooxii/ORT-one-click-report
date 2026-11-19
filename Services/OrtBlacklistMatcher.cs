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

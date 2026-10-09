using NUnit.Framework;
using ORT一键报告.Plans.ViewModels;
using System.Collections.Generic;

namespace ORT一键报告.Tests
{
    /// <summary>
    /// 搜索栏「最近搜索记忆」的规则测试：去重（不区分大小写）、最新在前、条数上限。
    /// </summary>
    [TestFixture]
    public class SearchHistoryRulesTests
    {
        [Test]
        public void 搜索记忆_最新在前并去重()
        {
            List<string> history = [];
            PlansViewModel.RememberSearchTerm(history, "WAQ");
            PlansViewModel.RememberSearchTerm(history, "DAQ");
            PlansViewModel.RememberSearchTerm(history, "waq");   // 大小写不同视为同一条

            Assert.That(history, Is.EqualTo(new[] { "waq", "DAQ" }));
        }

        [Test]
        public void 搜索记忆_首尾空白忽略_空关键字不记()
        {
            List<string> history = [];
            PlansViewModel.RememberSearchTerm(history, "  RT2609  ");
            PlansViewModel.RememberSearchTerm(history, "   ");
            PlansViewModel.RememberSearchTerm(history, null);

            Assert.That(history, Is.EqualTo(new[] { "RT2609" }));
        }

        [Test]
        public void 搜索记忆_超出上限时丢掉最旧的()
        {
            List<string> history = [];
            for (int i = 1; i <= 12; i++)
            {
                PlansViewModel.RememberSearchTerm(history, $"KW{i:D2}");
            }

            Assert.That(history.Count, Is.EqualTo(PlansViewModel.MaxSearchHistory));
            Assert.That(history[0], Is.EqualTo("KW12"));
            Assert.That(history[history.Count - 1], Is.EqualTo("KW03"));
        }

        [Test]
        public void 搜索记忆_重复搜索会提到最前()
        {
            List<string> history = ["A", "B", "C"];
            PlansViewModel.RememberSearchTerm(history, "C");

            Assert.That(history, Is.EqualTo(new[] { "C", "A", "B" }));
        }
    }
}

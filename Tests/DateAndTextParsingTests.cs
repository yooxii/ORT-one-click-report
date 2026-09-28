using NUnit.Framework;
using ORT一键报告.Models;
using ORT一键报告.Services;
using ORT一键报告.Utils;
using System;

namespace ORT一键报告.Tests
{
    /// <summary>
    /// 日期解析（PlanExcelService.ParseAnyDate）与文本归一化（OrtPlanParser / 阶段 / 测试种类）测试：
    /// 日期解析为历史高发问题区（"8月7日"、短横线、带引号格式等）。
    /// </summary>
    [TestFixture]
    public class DateAndTextParsingTests
    {
        /* ###############################  日期解析  ################################ */

        [TestCase("2026/8/18")]
        [TestCase("2026-8-18")]
        [TestCase("2026/08/18")]
        [TestCase(" 2026/8/18 ")]
        public void 数字日期格式_正确解析(string text)
            => Assert.That(PlanExcelService.ParseAnyDate(text)?.Date, Is.EqualTo(new DateTime(2026, 8, 18)));

        [Test]
        public void 带时间的日期_解析出日期部分()
            => Assert.That(PlanExcelService.ParseAnyDate("2026/8/18 10:30")?.Date, Is.EqualTo(new DateTime(2026, 8, 18)));

        [Test]
        public void 中文月日_使用回退年份()
            => Assert.That(PlanExcelService.ParseAnyDate("8月7日", 2026), Is.EqualTo(new DateTime(2026, 8, 7)));

        [Test]
        public void 中文月日_带转义引号也能解析()
            => Assert.That(PlanExcelService.ParseAnyDate("1\"月\"9\"日\"", 2026), Is.EqualTo(new DateTime(2026, 1, 9)));

        [Test]
        public void 中文月日_无回退年份_使用当前年()
            => Assert.That(PlanExcelService.ParseAnyDate("8月7日")?.Year, Is.EqualTo(DateTime.Now.Year));

        [TestCase(null)]
        [TestCase("")]
        [TestCase("   ")]
        [TestCase("not-a-date")]
        public void 无效文本_返回空(string text)
            => Assert.That(PlanExcelService.ParseAnyDate(text), Is.Null);

        [Test]
        public void 中文月日_非法月日_返回空()
            => Assert.That(PlanExcelService.ParseAnyDate("13月40日", 2026), Is.Null);

        /* ###############################  ORT Plan 文本归一化  ################################ */

        [Test]
        public void 归一化_压缩行内空白并去掉空行()
            => Assert.That(OrtPlanParser.Normalize(" a   b \r\n\r\n c "), Is.EqualTo("a b\nc"));

        [Test]
        public void 归一化_不换行空格与摄氏度统一()
            => Assert.That(OrtPlanParser.Normalize("25\u00A0℃"), Is.EqualTo("25 °C"));

        [TestCase(null, "")]
        [TestCase("", "")]
        [TestCase("   ", "")]
        public void 归一化_空白输入_返回空串(string input, string expected)
            => Assert.That(OrtPlanParser.Normalize(input), Is.EqualTo(expected));

        [Test]
        public void 名称归一化_去全部空白并大写()
            => Assert.That(OrtPlanParser.NormalizeName(" ab cd 亿 1 "), Is.EqualTo("ABCD亿1"));

        [Test]
        public void 名称归一化_空输入_返回空串()
            => Assert.That(OrtPlanParser.NormalizeName(null), Is.EqualTo(""));

        /* ###############################  阶段归并  ################################ */

        [TestCase("EVT", PlanStage.NPI)]
        [TestCase("DVT", PlanStage.NPI)]
        [TestCase("pvt", PlanStage.NPI)]
        [TestCase("Proto", PlanStage.NPI)]
        [TestCase("NPI", PlanStage.NPI)]
        [TestCase("新機種", PlanStage.NPI)]
        [TestCase("MP", PlanStage.MP)]
        [TestCase("Mass Production", PlanStage.MP)]
        [TestCase("", PlanStage.MP)]
        [TestCase(null, PlanStage.MP)]
        public void 阶段归并_试产按NPI_其余按MP(string raw, string expected)
            => Assert.That(PlanStage.Normalize(raw), Is.EqualTo(expected));

        [Test]
        public void 阶段显示名()
        {
            Assert.That(PlanStage.Display(PlanStage.NPI), Is.EqualTo("新机种"));
            Assert.That(PlanStage.Display(PlanStage.MP), Is.EqualTo("量产"));
        }

        /* ###############################  测试种类归并  ################################ */

        [Test]
        public void 测试种类_EMC与EMI归为同类()
        {
            string emc = TestCategories.Normalize("EMC");
            Assert.That(emc, Is.Not.Null);
            Assert.That(TestCategories.Normalize("EMI"), Is.EqualTo(emc));
        }

        [Test]
        public void 测试种类_可靠性类关键词归为同类()
        {
            string reliability = TestCategories.Normalize("RELIABILITY TEST");
            Assert.That(reliability, Is.Not.Null);
            Assert.That(TestCategories.Normalize("Environment Tests"), Is.EqualTo(reliability));
            Assert.That(TestCategories.Normalize("Mechanical"), Is.EqualTo(reliability));
            Assert.That(TestCategories.Normalize("Safety"), Is.EqualTo(reliability));
        }

        [Test]
        public void 测试种类_认不出返回空()
        {
            Assert.That(TestCategories.Normalize("别的写法"), Is.Null);
            Assert.That(TestCategories.Normalize((string)null), Is.Null);
        }

        [Test]
        public void 测试种类_多个历史分类_取出现次数最多的()
        {
            string expected = TestCategories.Normalize("EMC");
            Assert.That(TestCategories.Normalize(["EMC", "EMC", "RELIABILITY TEST"]), Is.EqualTo(expected));
            Assert.That(TestCategories.Normalize([]), Is.Null);
        }
    }
}

using NUnit.Framework;
using ORT一键报告.Plans.ViewModels;
using System;

namespace ORT一键报告.Tests
{
    /// <summary>
    /// 回线RT工令编号体检的纯规则测试：
    /// 空值算正常、编号年月与领用日期或回线日期任一对上就算正常（历史数据两种口径都有）、
    /// 两个都对不上才报月段问题并给出建议编号（保留原序号位数）、
    /// 不是「RTAH + 四位年月 + 序号」形状的算格式异常。
    /// </summary>
    [TestFixture]
    public class ReturnRtCodeCheckTests
    {
        private static readonly DateTime January9 = new(2026, 1, 9);
        private static readonly DateTime January20 = new(2026, 1, 20);
        private static readonly DateTime March27 = new(2026, 3, 27);
        private static readonly DateTime April13 = new(2026, 4, 13);
        private static readonly DateTime October9 = new(2026, 10, 9);

        /* ###############################  年月提取  ################################ */

        [TestCase("RTAH2610001", "2610")]
        [TestCase("rtah2601013", "2601")]
        [TestCase("  RTAH260301  ", "2603")]
        public void 取年月段_正常编号(string code, string expected)
            => Assert.That(RequisitionEditRules.ReturnRtYearMonth(code), Is.EqualTo(expected));

        [TestCase(null)]
        [TestCase("")]
        [TestCase("   ")]
        [TestCase("RTAH")]
        [TestCase("RTAH261")]
        [TestCase("RT260901")]
        [TestCase("濕熱處理測試 報廢處理不回綫")]
        public void 取年月段_非编号形状返回空(string code)
            => Assert.That(RequisitionEditRules.ReturnRtYearMonth(code), Is.Null);

        /* ###############################  体检判定  ################################ */

        [TestCase(null)]
        [TestCase("")]
        [TestCase("   ")]
        public void 体检_空值算正常_未登记回线(string code)
            => Assert.That(RequisitionEditRules.InspectReturnRt(code, January9, null), Is.Null);

        [Test]
        public void 体检_年月与领用日期一致算正常()
            => Assert.That(RequisitionEditRules.InspectReturnRt("RTAH2601013", January9, null), Is.Null);

        [Test]
        public void 体检_年月与回线日期一致也算正常_历史数据有按回线日期编号的()
        {
            // 3 月领用、4 月回线 → 编号是 4 月的号，属于正常口径，不该报异常
            Assert.That(RequisitionEditRules.InspectReturnRt("RTAH260401", March27, April13), Is.Null);
        }

        [Test]
        public void 体检_年月不符_两个日期都对不上才报问题()
        {
            (ReturnRtCodeIssueKind Kind, string SuggestedCode)? found =
                RequisitionEditRules.InspectReturnRt("RTAH2610002", January9, January20);

            Assert.That(found, Is.Not.Null);
            Assert.That(found.Value.Kind, Is.EqualTo(ReturnRtCodeIssueKind.MonthMismatch));
            // 序号位数照旧保留（三位 002 → 仍三位），按领用日期的年月给建议
            Assert.That(found.Value.SuggestedCode, Is.EqualTo("RTAH2601002"));
        }

        [Test]
        public void 体检_年月不符_两位序号同样保留()
        {
            (ReturnRtCodeIssueKind Kind, string SuggestedCode)? found =
                RequisitionEditRules.InspectReturnRt("RTAH260305", October9, null);

            Assert.That(found.Value.Kind, Is.EqualTo(ReturnRtCodeIssueKind.MonthMismatch));
            Assert.That(found.Value.SuggestedCode, Is.EqualTo("RTAH261005"));
        }

        [Test]
        public void 体检_十月记录配十月日期算正常()
            => Assert.That(RequisitionEditRules.InspectReturnRt("RTAH261005", October9, null), Is.Null);

        [Test]
        public void 体检_没有领用日期时按回线日期判断()
        {
            // 只填了回线日期：与之一致就算正常
            Assert.That(RequisitionEditRules.InspectReturnRt("RTAH260401", null, April13), Is.Null);
            // 与回线日期也不符 → 报问题，建议编号用回线日期的年月
            (ReturnRtCodeIssueKind Kind, string SuggestedCode)? found =
                RequisitionEditRules.InspectReturnRt("RTAH2610001", null, January20);
            Assert.That(found.Value.Kind, Is.EqualTo(ReturnRtCodeIssueKind.MonthMismatch));
            Assert.That(found.Value.SuggestedCode, Is.EqualTo("RTAH2601001"));
        }

        [Test]
        public void 体检_两个日期都没有时只做格式检查()
        {
            // 形状正确但无从判断年月 → 不算问题
            Assert.That(RequisitionEditRules.InspectReturnRt("RTAH2610002", null, null), Is.Null);
            // 形状不对 → 仍然报格式异常
            (ReturnRtCodeIssueKind Kind, string SuggestedCode)? found =
                RequisitionEditRules.InspectReturnRt("回綫", null, null);
            Assert.That(found.Value.Kind, Is.EqualTo(ReturnRtCodeIssueKind.UnrecognizedFormat));
        }

        [Test]
        public void 体检_不是编号形状_报格式异常且无建议编号()
        {
            (ReturnRtCodeIssueKind Kind, string SuggestedCode)? found =
                RequisitionEditRules.InspectReturnRt("濕熱處理測試 報廢處理不回綫", new DateTime(2026, 5, 7), null);

            Assert.That(found, Is.Not.Null);
            Assert.That(found.Value.Kind, Is.EqualTo(ReturnRtCodeIssueKind.UnrecognizedFormat));
            Assert.That(found.Value.SuggestedCode, Is.Null);
        }

        [Test]
        public void 体检_大小写与首尾空白不影响判定()
            => Assert.That(RequisitionEditRules.InspectReturnRt("  rtah2601013  ", January9, null), Is.Null);
    }
}
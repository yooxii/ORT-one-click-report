using NUnit.Framework;
using ORT一键报告.Models;
using ORT一键报告.Plans.ViewModels;

namespace ORT一键报告.Tests
{
    /// <summary>
    /// 领退表新增/编辑窗口的纯规则测试：
    /// 单体去向按机种名称判定（开头为 W → 报废，其余 → 入库）、
    /// 完工令（Work Order）自动补全线别与 D/C（界面上显示为周期），
    /// 以及「保存并继续」时编号的递增避重。
    /// </summary>
    [TestFixture]
    public class RequisitionEditRulesTests
    {
        /* ###############################  单体去向  ################################ */

        [Test]
        public void 单体去向_机种以W开头判为报废()
            => Assert.That(RequisitionEditRules.DispositionFromModel("WAQ2601001"),
                Is.EqualTo(RequisitionDispositionKind.Scrap));

        [Test]
        public void 单体去向_机种不以W开头判为入库()
            => Assert.That(RequisitionEditRules.DispositionFromModel("DAQ2601001"),
                Is.EqualTo(RequisitionDispositionKind.StockIn));

        [TestCase("w1234567")]
        [TestCase("  WaQ1234  ")]
        public void 单体去向_大小写与首尾空白不影响判定(string model)
            => Assert.That(RequisitionEditRules.DispositionFromModel(model),
                Is.EqualTo(RequisitionDispositionKind.Scrap));

        [TestCase(null)]
        [TestCase("")]
        [TestCase("   ")]
        public void 单体去向_机种为空时默认入库(string model)
            => Assert.That(RequisitionEditRules.DispositionFromModel(model),
                Is.EqualTo(RequisitionDispositionKind.StockIn));

        /* ###############################  完工令补全  ################################ */

        [Test]
        public void 完工令_线别取倒数第六位起三位_DC取年份后两位加周别()
        {
            const string workOrder = "WK2601THRSA1";   // 下标 6..8=THR，下标 9..10=SA
            Assert.That(RequisitionEditRules.LineNoFromWorkOrder(workOrder), Is.EqualTo("THR"));
            // D/C（周期）= 年份后两位 + 倒数第三位起的两位（末尾一位是校验位，不取）
            Assert.That(RequisitionEditRules.DcFromWorkOrder(workOrder, 2026), Is.EqualTo("26SA"));
            Assert.That(RequisitionEditRules.DcFromWorkOrder("WK2601THR331", 2026), Is.EqualTo("2633"));
            Assert.That(RequisitionEditRules.DcFromWorkOrder("WK2601THR331", 2031), Is.EqualTo("3133"));
        }

        [Test]
        public void 完工令_首尾空白不影响解析()
        {
            Assert.That(RequisitionEditRules.LineNoFromWorkOrder("  WK2601THRSA1 "), Is.EqualTo("THR"));
            Assert.That(RequisitionEditRules.DcFromWorkOrder("  WK2601THR331 ", 2026), Is.EqualTo("2633"));
        }

        [Test]
        public void 完工令_太短时取不到补全值()
        {
            Assert.That(RequisitionEditRules.LineNoFromWorkOrder("12345"), Is.Null);
            Assert.That(RequisitionEditRules.DcFromWorkOrder("12", 2026), Is.Null);
            Assert.That(RequisitionEditRules.DcFromWorkOrder("123", 2026), Is.EqualTo("2612"));
        }

        [TestCase(null)]
        [TestCase("")]
        public void 完工令_为空时取不到补全值(string workOrder)
        {
            Assert.That(RequisitionEditRules.LineNoFromWorkOrder(workOrder), Is.Null);
            Assert.That(RequisitionEditRules.DcFromWorkOrder(workOrder, 2026), Is.Null);
        }

        /* ###############################  机种名称补全  ################################ */

        private static readonly string[] Models = ["WAQ2601001", "WAQ26010012", "DAQ2601001"];

        [Test]
        public void 机种补全_取最短的以输入开头的已有名称()
        {
            Assert.That(RequisitionEditRules.ModelSuggestion(Models, "WAQ"), Is.EqualTo("WAQ2601001"));
            Assert.That(RequisitionEditRules.ModelSuggestion(Models, "WAQ2601001"), Is.EqualTo("WAQ26010012"));
            Assert.That(RequisitionEditRules.ModelSuggestion(Models, "waq2601"), Is.EqualTo("WAQ2601001"));   // 不区分大小写
            Assert.That(RequisitionEditRules.ModelSuggestion(Models, " WAQ2601 "), Is.EqualTo("WAQ2601001")); // 首尾空白忽略
        }

        [Test]
        public void 机种补全_没有更长候选时不提示()
        {
            Assert.That(RequisitionEditRules.ModelSuggestion(Models, "WAQ26010012"), Is.Null);  // 已完整
            Assert.That(RequisitionEditRules.ModelSuggestion(Models, "XYZ"), Is.Null);          // 没有匹配
            Assert.That(RequisitionEditRules.ModelSuggestion(Models, ""), Is.Null);             // 没输入
            Assert.That(RequisitionEditRules.ModelSuggestion(null, "WAQ"), Is.Null);            // 没有候选
        }

        /* ###############################  编号递增  ################################ */

        [Test]
        public void 工作编号_序号加一并保持两位()
        {
            Assert.That(RequisitionEditRules.NextJobNo("RT260901"), Is.EqualTo("RT260902"));
            Assert.That(RequisitionEditRules.NextJobNo("QRT260909"), Is.EqualTo("QRT260910"));
        }

        [Test]
        public void 工作编号_序号跨过两位上限时按实际位数展开()
            => Assert.That(RequisitionEditRules.NextJobNo("RT260999"), Is.EqualTo("RT2609100"));

        [Test]
        public void 回线RT工令_序号加一()
            => Assert.That(RequisitionEditRules.NextReturnRtOrder("RTAH260909"), Is.EqualTo("RTAH260910"));

        [TestCase(null)]
        [TestCase("")]
        [TestCase("RT2609")]
        [TestCase("随便写")]
        public void 编号_格式认不出时返回空(string value)
        {
            Assert.That(RequisitionEditRules.NextJobNo(value), Is.Null);
            Assert.That(RequisitionEditRules.NextReturnRtOrder(value), Is.Null);
        }
    }
}

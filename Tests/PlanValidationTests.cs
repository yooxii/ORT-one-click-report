using NUnit.Framework;
using ORT一键报告.Utils;

namespace ORT一键报告.Tests
{
    /// <summary>
    /// 计划数据编辑校验（PlanValidation）与状况归类（PlanStatusKind）测试：
    /// 覆盖工作编号格式、回线RT工令格式、状况枚举、字典约束与三种状况归类的容错写法。
    /// </summary>
    [TestFixture]
    public class PlanValidationTests
    {
        /* ###############################  工作编号  ################################ */

        [TestCase("RT260801")]
        [TestCase("QRT260901")]
        [TestCase("rt260801")]      // 前缀不区分大小写
        [TestCase("Qrt261201")]
        [TestCase("RT26080100")]    // 编号超过两位也允许（MAX+1 展开）
        public void 合法工作编号_校验通过(string jobNo)
            => Assert.That(PlanValidation.ValidateJobNo(jobNo), Is.Null);

        [TestCase(null)]
        [TestCase("")]
        [TestCase("   ")]
        public void 空工作编号_允许为空(string jobNo)
            => Assert.That(PlanValidation.ValidateJobNo(jobNo), Is.Null);

        [TestCase("RT2608")]        // 缺末尾编号
        [TestCase("RT2608A")]       // 编号非数字
        [TestCase("XX260801")]      // 前缀不是 RT/QRT
        [TestCase("RT26801")]       // 年月只有 3 位
        public void 非法工作编号_给出格式说明(string jobNo)
        {
            string error = PlanValidation.ValidateJobNo(jobNo);
            Assert.That(error, Is.Not.Null);
            Assert.That(error, Does.Contain("QRT/RT"));
        }

        [Test]
        public void 编号为00_提示从01开始()
        {
            string error = PlanValidation.ValidateJobNo("RT260800");
            Assert.That(error, Does.Contain("01"));
        }

        /* ###############################  回线RT工令  ################################ */

        [TestCase("RTAH260901")]
        [TestCase("rtah260901")]      // 前缀不区分大小写
        [TestCase("RTAH261299")]
        [TestCase("RTAH2609100")]     // 当月超过 99 笔时生成器展开为 3 位，同样允许
        [TestCase(" RTAH260901 ")]    // 首尾空白容忍
        public void 合法回线RT工令_校验通过(string order)
            => Assert.That(PlanValidation.ValidateReturnRtOrder(order), Is.Null);

        [TestCase(null)]
        [TestCase("")]
        [TestCase("   ")]
        public void 空回线RT工令_允许为空(string order)
            => Assert.That(PlanValidation.ValidateReturnRtOrder(order), Is.Null);

        [TestCase("RT260901")]        // 前缀不是 RTAH
        [TestCase("XXAH260901")]      // 前缀错误
        [TestCase("RTAH2609")]        // 缺末尾编号
        [TestCase("RTAH26090")]       // 编号只有 1 位
        [TestCase("RTAH26091")]       // 年月+编号共 5 位，凑不出至少 2 位编号
        [TestCase("RTAH2609A1")]      // 编号含字母
        [TestCase("RTAH 260901")]     // 中间有空格
        public void 非法回线RT工令_给出格式说明(string order)
        {
            string error = PlanValidation.ValidateReturnRtOrder(order);
            Assert.That(error, Is.Not.Null);
            Assert.That(error, Does.Contain("RTAH"));
        }

        [Test]
        public void 回线RT工令编号为00_提示从01开始()
        {
            string error = PlanValidation.ValidateReturnRtOrder("RTAH260900");
            Assert.That(error, Does.Contain("01"));
        }

        /* ###############################  状况  ################################ */

        [TestCase("Ongoing")]
        [TestCase("close")]
        [TestCase("PENDING")]
        [TestCase(" Close ")]
        [TestCase(null)]
        [TestCase("")]
        public void 合法状况_校验通过(string status)
            => Assert.That(PlanValidation.ValidateStatus(status), Is.Null);

        [Test]
        public void 非法状况_给出枚举说明()
        {
            string error = PlanValidation.ValidateStatus("Closed");   // 历史写法不算合法枚举（归类时会算已结案）
            Assert.That(error, Is.Not.Null);
            Assert.That(error, Does.Contain("Ongoing"));
        }

        /* ###############################  字典约束  ################################ */

        [Test]
        public void 字典内的值_校验通过()
            => Assert.That(PlanValidation.ValidateInCatalog("TestA", ["TestA", "TestB"], "测试项目"), Is.Null);

        [Test]
        public void 字典外的值_提示先到管理模块添加()
        {
            string error = PlanValidation.ValidateInCatalog("TestC", ["TestA"], "测试项目");
            Assert.That(error, Is.Not.Null);
            Assert.That(error, Does.Contain("测试项目"));
            Assert.That(error, Does.Contain("TestC"));
        }

        [Test]
        public void 空值不校验字典()
            => Assert.That(PlanValidation.ValidateInCatalog("  ", ["TestA"], "测试项目"), Is.Null);

        /* ###############################  状况归类  ################################ */

        [TestCase("Ongoing", PlanStatusKind.Ongoing)]
        [TestCase("进行中", PlanStatusKind.Ongoing)]
        [TestCase("測試中", PlanStatusKind.Ongoing)]
        [TestCase("Pending", PlanStatusKind.Pending)]
        [TestCase("预排测试", PlanStatusKind.Pending)]
        [TestCase("Close", PlanStatusKind.Closed)]
        [TestCase("Closed", PlanStatusKind.Closed)]
        [TestCase("结案", PlanStatusKind.Closed)]
        [TestCase("已完成", PlanStatusKind.Closed)]
        [TestCase("  close  ", PlanStatusKind.Closed)]
        public void 状况归类_容错各种写法(string raw, string expected)
            => Assert.That(PlanStatusKind.Of(raw), Is.EqualTo(expected));

        [TestCase(null)]
        [TestCase("")]
        [TestCase("   ")]
        [TestCase("自定义写法")]
        public void 认不出的状况_归类为空(string raw)
            => Assert.That(PlanStatusKind.Of(raw), Is.Empty);
    }
}

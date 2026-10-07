using NUnit.Framework;
using ORT一键报告.Models;
using ORT一键报告.Plans.ViewModels;
using ORT一键报告.Services;
using System;

namespace ORT一键报告.Tests
{
    /// <summary>
    /// 领退表批量登记规则（RequisitionBatchRules）测试：多选批量回线/入库时，
    /// 哪些记录可执行、哪些必须跳过，判定必须与单条右键动作的分支限制一致。
    /// 期望文本统一用 LanguageService.Get 取（与实际界面同一来源），避免依赖运行环境语言。
    /// </summary>
    [TestFixture]
    public class RequisitionBatchRulesTests
    {
        private static string L(string key) => LanguageService.Get(key);

        private static Requisition Req(string disposition) => new()
        {
            Id = 1,
            RequisitionNo = "WL-001",
            ModelName = "M1",
            Disposition = disposition,
        };

        /* ###############################  批量回线  ################################ */

        [Test]
        public void 批量回线_入库去向_可执行()
        {
            Assert.That(RequisitionBatchRules.ReturnBlockReason(Req(RequisitionDispositionKind.StockIn)), Is.Null);
        }

        [Test]
        public void 批量回线_报废去向_跳过并给出原因()
        {
            Assert.That(RequisitionBatchRules.ReturnBlockReason(Req(RequisitionDispositionKind.Scrap)),
                Is.EqualTo(L("Msg_OpBlockedDispositionScrap")));
        }

        [Test]
        public void 批量回线_空记录_提示先选择记录()
        {
            Assert.That(RequisitionBatchRules.ReturnBlockReason(null), Is.EqualTo(L("Plans_Msg_SelectRequisition")));
        }

        /* ###############################  批量入库  ################################ */

        [Test]
        public void 批量入库_已回线的入库去向_可执行()
        {
            Requisition req = Req(RequisitionDispositionKind.StockIn);
            req.ReturnDate = new DateTime(2026, 10, 1);

            Assert.That(RequisitionBatchRules.StockInBlockReason(req), Is.Null);
        }

        [Test]
        public void 批量入库_未回线_跳过()
        {
            Assert.That(RequisitionBatchRules.StockInBlockReason(Req(RequisitionDispositionKind.StockIn)),
                Is.EqualTo(L("Msg_StockInNeedReturn")));
        }

        [Test]
        public void 批量入库_已报废_跳过()
        {
            Requisition req = Req(RequisitionDispositionKind.StockIn);
            req.ReturnDate = new DateTime(2026, 10, 1);
            req.ScrapQty = "3";

            Assert.That(RequisitionBatchRules.StockInBlockReason(req), Is.EqualTo(L("Msg_StockInAfterScrap")));
        }

        [Test]
        public void 批量入库_报废去向_跳过()
        {
            Assert.That(RequisitionBatchRules.StockInBlockReason(Req(RequisitionDispositionKind.Scrap)),
                Is.EqualTo(L("Msg_OpBlockedDispositionScrap")));
        }

        /* ###############################  批量删除  ################################ */

        [Test]
        public void 批量删除_任何记录都可标记删除()
        {
            Assert.That(RequisitionBatchRules.BlockReason(RequisitionBatchMode.Delete, Req(RequisitionDispositionKind.Scrap)), Is.Null);
            Assert.That(RequisitionBatchRules.BlockReason(RequisitionBatchMode.Delete, Req(RequisitionDispositionKind.StockIn)), Is.Null);
        }

        [Test]
        public void 按类型取原因_回线与入库各自的限制生效()
        {
            Requisition scrap = Req(RequisitionDispositionKind.Scrap);
            Requisition stockIn = Req(RequisitionDispositionKind.StockIn);

            Assert.That(RequisitionBatchRules.BlockReason(RequisitionBatchMode.Return, scrap), Is.EqualTo(L("Msg_OpBlockedDispositionScrap")));
            Assert.That(RequisitionBatchRules.BlockReason(RequisitionBatchMode.Return, stockIn), Is.Null);
            Assert.That(RequisitionBatchRules.BlockReason(RequisitionBatchMode.StockIn, stockIn), Is.EqualTo(L("Msg_StockInNeedReturn")));
        }
    }
}

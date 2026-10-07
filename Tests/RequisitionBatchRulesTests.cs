using NUnit.Framework;
using ORT一键报告.Models;
using ORT一键报告.Plans.ViewModels;
using ORT一键报告.Services;
using System;

namespace ORT一键报告.Tests
{
    /// <summary>
    /// 领退表登记规则（RequisitionBatchRules）测试：批量回线/删除哪些记录可执行，
    /// 以及「筛选待入库」的筛选口径与逐条登记「下一个待入库」的跳转口径，
    /// 判定必须与单条右键动作的分支限制一致。
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

        /* ###############################  入库登记的分支限制  ################################ */

        [Test]
        public void 入库_已回线的入库去向_可执行()
        {
            Requisition req = Req(RequisitionDispositionKind.StockIn);
            req.ReturnDate = new DateTime(2026, 10, 1);

            Assert.That(RequisitionBatchRules.StockInBlockReason(req), Is.Null);
        }

        [Test]
        public void 入库_未回线_跳过()
        {
            Assert.That(RequisitionBatchRules.StockInBlockReason(Req(RequisitionDispositionKind.StockIn)),
                Is.EqualTo(L("Msg_StockInNeedReturn")));
        }

        [Test]
        public void 入库_已报废_跳过()
        {
            Requisition req = Req(RequisitionDispositionKind.StockIn);
            req.ReturnDate = new DateTime(2026, 10, 1);
            req.ScrapQty = "3";

            Assert.That(RequisitionBatchRules.StockInBlockReason(req), Is.EqualTo(L("Msg_StockInAfterScrap")));
        }

        [Test]
        public void 入库_报废去向_跳过()
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
        public void 按类型取原因_回线限制生效()
        {
            Requisition scrap = Req(RequisitionDispositionKind.Scrap);
            Requisition stockIn = Req(RequisitionDispositionKind.StockIn);

            Assert.That(RequisitionBatchRules.BlockReason(RequisitionBatchMode.Return, scrap), Is.EqualTo(L("Msg_OpBlockedDispositionScrap")));
            Assert.That(RequisitionBatchRules.BlockReason(RequisitionBatchMode.Return, stockIn), Is.Null);
        }

        /* ###############################  表格筛选口径  ################################ */

        [Test]
        public void 待回线_入库去向未回线_命中()
        {
            Assert.That(RequisitionBatchRules.NeedsReturn(Req(RequisitionDispositionKind.StockIn)), Is.True);
        }

        [Test]
        public void 待回线_报废去向或已回线_不命中()
        {
            Requisition returned = Req(RequisitionDispositionKind.StockIn);
            returned.ReturnDate = new DateTime(2026, 10, 1);

            Assert.That(RequisitionBatchRules.NeedsReturn(Req(RequisitionDispositionKind.Scrap)), Is.False);
            Assert.That(RequisitionBatchRules.NeedsReturn(returned), Is.False);
        }

        [Test]
        public void 筛选待入库_未报废且未入库_命中_含未回线的()
        {
            Requisition notReturned = Req(RequisitionDispositionKind.StockIn);
            Requisition returned = Req(RequisitionDispositionKind.StockIn);
            returned.ReturnDate = new DateTime(2026, 10, 1);

            Assert.That(RequisitionBatchRules.NeedsStockIn(notReturned), Is.True);
            Assert.That(RequisitionBatchRules.NeedsStockIn(returned), Is.True);
        }

        [Test]
        public void 筛选待入库_已入库或报废_不命中()
        {
            Requisition stockedByNo = Req(RequisitionDispositionKind.StockIn);
            stockedByNo.StockInNo = "RK-001";
            Requisition stockedByDate = Req(RequisitionDispositionKind.StockIn);
            stockedByDate.StockInDate = new DateTime(2026, 10, 6);
            Requisition scrapped = Req(RequisitionDispositionKind.StockIn);
            scrapped.ScrapQty = "3";

            Assert.That(RequisitionBatchRules.NeedsStockIn(stockedByNo), Is.False);
            Assert.That(RequisitionBatchRules.NeedsStockIn(stockedByDate), Is.False);
            Assert.That(RequisitionBatchRules.NeedsStockIn(scrapped), Is.False);
            Assert.That(RequisitionBatchRules.NeedsStockIn(Req(RequisitionDispositionKind.Scrap)), Is.False);
        }

        [Test]
        public void 筛选口径_按批量类型取_删除不过滤()
        {
            Requisition scrap = Req(RequisitionDispositionKind.Scrap);

            Assert.That(RequisitionBatchRules.MatchesQuickFilter(RequisitionBatchMode.Return, scrap), Is.False);
            Assert.That(RequisitionBatchRules.MatchesQuickFilter(RequisitionBatchMode.Delete, scrap), Is.True);
        }

        /* ###############################  「下一个待入库」的跳转口径  ################################ */

        [Test]
        public void 下一个待入库_已回线且未入库_可作为跳转目标()
        {
            Requisition req = Req(RequisitionDispositionKind.StockIn);
            req.ReturnDate = new DateTime(2026, 10, 1);

            Assert.That(RequisitionBatchRules.CanRegisterStockIn(req), Is.True);
        }

        [Test]
        public void 下一个待入库_未回线已入库或报废_不作为跳转目标()
        {
            Requisition notReturned = Req(RequisitionDispositionKind.StockIn);
            Requisition stocked = Req(RequisitionDispositionKind.StockIn);
            stocked.ReturnDate = new DateTime(2026, 10, 1);
            stocked.StockInNo = "RK-001";
            Requisition scrapped = Req(RequisitionDispositionKind.StockIn);
            scrapped.ReturnDate = new DateTime(2026, 10, 1);
            scrapped.ScrapQty = "1";

            Assert.That(RequisitionBatchRules.CanRegisterStockIn(notReturned), Is.False);
            Assert.That(RequisitionBatchRules.CanRegisterStockIn(stocked), Is.False);
            Assert.That(RequisitionBatchRules.CanRegisterStockIn(scrapped), Is.False);
            Assert.That(RequisitionBatchRules.CanRegisterStockIn(Req(RequisitionDispositionKind.Scrap)), Is.False);
            Assert.That(RequisitionBatchRules.CanRegisterStockIn(null), Is.False);
        }
    }
}

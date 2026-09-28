using NUnit.Framework;
using ORT一键报告.Models;
using ORT一键报告.Plans.ViewModels;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ORT一键报告.Tests
{
    /// <summary>
    /// 流程查看（FlowViewModel）测试：覆盖其他部门申请测试流程（QRT，含「单体归还」步骤）
    /// 与一般流程（领用，含「单体去向」计划分支、双分支异常提示）的状态推导。
    /// 期望文本统一用 LanguageService.Get 取（与实际界面同一来源），避免依赖运行环境语言。
    /// </summary>
    [TestFixture]
    public class FlowViewModelTests
    {
        private static string L(string key) => LanguageService.Get(key);

        private static Plan QrtPlan() => new() { JobNo = "QRT260901", ModelName = "M1", Status = "Ongoing" };

        private static Requisition Req(string disposition) => new()
        {
            Id = 1,
            RequisitionNo = "WL-001",
            ModelName = "M1",
            RequisitionDate = new DateTime(2026, 9, 1),
            Disposition = disposition,
        };

        private static FlowStepItem FindStep(FlowViewModel vm, string title)
        {
            IEnumerable<FlowStepItem> all = vm.Rows.SelectMany(r => r.Items)
                .Concat(vm.Branches.SelectMany(b => b.Steps));
            if (vm.CloseStep != null)
            {
                all = all.Append(vm.CloseStep);
            }
            return all.First(s => s.Title == title);
        }

        /* ###############################  其他部门申请测试流程（QRT）  ################################ */

        [Test]
        public void QRT_完成报告后未登记归还_归还步骤为进行中()
        {
            Plan plan = QrtPlan();
            plan.ReportStatus = ReportStatusKind.Complete;

            FlowStepItem step = FindStep(FlowViewModel.BuildExternalTest(plan), L("Flow_Step_UnitReturn"));

            Assert.That(step.State, Is.EqualTo(FlowState.Current));
            Assert.That(step.Evidence, Is.EqualTo(L("Flow_Ev_UnitReturnPending")));
        }

        [Test]
        public void QRT_报告进行中_归还步骤未到达()
        {
            Plan plan = QrtPlan();
            plan.ReportStatus = ReportStatusKind.InProgress;

            FlowStepItem step = FindStep(FlowViewModel.BuildExternalTest(plan), L("Flow_Step_UnitReturn"));

            Assert.That(step.State, Is.EqualTo(FlowState.Waiting));
        }

        [Test]
        public void QRT_无需报告且未登记归还_归还步骤为进行中()
        {
            Plan plan = QrtPlan();
            plan.ReportStatus = ReportStatusKind.NotRequired;

            FlowStepItem step = FindStep(FlowViewModel.BuildExternalTest(plan), L("Flow_Step_UnitReturn"));

            Assert.That(step.State, Is.EqualTo(FlowState.Current));
        }

        [Test]
        public void QRT_已登记归还日期_归还步骤完成并带日期依据()
        {
            Plan plan = QrtPlan();
            plan.ReportStatus = ReportStatusKind.Complete;
            plan.UnitReturnDate = new DateTime(2026, 9, 28);

            FlowStepItem step = FindStep(FlowViewModel.BuildExternalTest(plan), L("Flow_Step_UnitReturn"));

            Assert.That(step.State, Is.EqualTo(FlowState.Done));
            // 与 FlowViewModel 相同的格式化方式（当前文化 + yyyy/M/d），保证断言与实现一致
            Assert.That(step.Evidence,
                Is.EqualTo(string.Format(L("Flow_Ev_UnitReturnDoneFormat"), plan.UnitReturnDate)));
        }

        [Test]
        public void QRT_归还步骤位于完成报告之后_结案为共用节点()
        {
            Plan plan = QrtPlan();
            plan.ReportStatus = ReportStatusKind.Complete;

            FlowViewModel vm = FlowViewModel.BuildExternalTest(plan);
            List<string> titles = vm.Rows.SelectMany(r => r.Items).Select(s => s.Title).ToList();
            int reportIndex = titles.IndexOf(L("Flow_Step_ReportOptional"));
            int returnIndex = titles.IndexOf(L("Flow_Step_UnitReturn"));

            Assert.That(reportIndex, Is.GreaterThanOrEqualTo(0));
            Assert.That(returnIndex, Is.EqualTo(reportIndex + 1));
            Assert.That(vm.CloseStep, Is.Not.Null);
            Assert.That(vm.CloseStep.Title, Is.EqualTo(L("Flow_Step_Close")));
        }

        [Test]
        public void QRT_已结案但未登记归还_底部给出核对提示()
        {
            Plan plan = QrtPlan();
            plan.Status = "Close";
            plan.ReportStatus = ReportStatusKind.Complete;

            FlowViewModel vm = FlowViewModel.BuildExternalTest(plan);

            Assert.That(vm.FooterNote, Does.Contain(L("Flow_WarnClosedWithoutUnitReturn")));
            Assert.That(vm.CurrentText, Does.Contain(L("Flow_Step_UnitReturn")));
        }

        [Test]
        public void QRT_已结案且已登记归还_无核对提示()
        {
            Plan plan = QrtPlan();
            plan.Status = "Close";
            plan.ReportStatus = ReportStatusKind.Complete;
            plan.UnitReturnDate = new DateTime(2026, 9, 28);

            FlowViewModel vm = FlowViewModel.BuildExternalTest(plan);

            Assert.That(vm.FooterNote, Does.Not.Contain(L("Flow_WarnClosedWithoutUnitReturn")));
        }

        [Test]
        public void QRT_空计划_不抛异常且归还步骤未到达()
        {
            FlowViewModel vm = FlowViewModel.BuildExternalTest(null);

            FlowStepItem step = FindStep(vm, L("Flow_Step_UnitReturn"));
            Assert.That(step.State, Is.EqualTo(FlowState.Waiting));
            Assert.That(step.Evidence, Is.EqualTo(L("Flow_Ev_NoPlan")));
        }

        /* ###############################  一般流程（领用）：去向与分支  ################################ */

        [Test]
        public void 一般流程_去向为报废且操作未登记_报废分支显示计划走此分支()
        {
            FlowViewModel vm = FlowViewModel.BuildGeneral(Req(RequisitionDispositionKind.Scrap), null, null, null);

            Assert.That(vm.Branches[0].StateText, Is.EqualTo(L("Flow_Branch_Planned")));
            Assert.That(vm.Branches[1].StateText, Is.EqualTo(L("Flow_Branch_Undecided")));
            Assert.That(vm.Branches[0].IsActive, Is.False);
            Assert.That(vm.Branches[1].IsDimmed, Is.False);
        }

        [Test]
        public void 一般流程_去向为入库且操作未登记_入库分支显示计划走此分支()
        {
            FlowViewModel vm = FlowViewModel.BuildGeneral(Req(RequisitionDispositionKind.StockIn), null, null, null);

            Assert.That(vm.Branches[1].StateText, Is.EqualTo(L("Flow_Branch_Planned")));
            Assert.That(vm.Branches[0].StateText, Is.EqualTo(L("Flow_Branch_Undecided")));
        }

        [Test]
        public void 一般流程_已登记报废_报废分支为走此分支且入库分支置灰()
        {
            Requisition req = Req(RequisitionDispositionKind.Scrap);
            req.ScrapQty = "10";
            req.ScrapDate = new DateTime(2026, 9, 20);

            FlowViewModel vm = FlowViewModel.BuildGeneral(req, null, null, null);

            Assert.That(vm.Branches[0].StateText, Is.EqualTo(L("Flow_Branch_Taken")));
            Assert.That(vm.Branches[0].IsActive, Is.True);
            Assert.That(vm.Branches[1].StateText, Is.EqualTo(L("Flow_Branch_NotTaken")));
            Assert.That(vm.Branches[1].IsDimmed, Is.True);
        }

        [Test]
        public void 一般流程_去向与已走分支不一致_以实际操作登记为准()
        {
            // 去向填了「入库」，但已经登记了报废：分支判定优先按操作数据
            Requisition req = Req(RequisitionDispositionKind.StockIn);
            req.ScrapQty = "2";

            FlowViewModel vm = FlowViewModel.BuildGeneral(req, null, null, null);

            Assert.That(vm.Branches[0].StateText, Is.EqualTo(L("Flow_Branch_Taken")));
            Assert.That(vm.Branches[1].StateText, Is.EqualTo(L("Flow_Branch_NotTaken")));
        }

        [Test]
        public void 一般流程_两条分支都有登记_底部出现异常提示()
        {
            Requisition req = Req(null);
            req.ScrapQty = "1";
            req.StockInNo = "RK-1";

            FlowViewModel vm = FlowViewModel.BuildGeneral(req, null, null, null);

            Assert.That(vm.FooterNote, Does.StartWith(L("Flow_WarnBothBranches")));
        }

        [Test]
        public void 一般流程_无领用记录_分支待定且不抛异常()
        {
            FlowViewModel vm = FlowViewModel.BuildGeneral(null, null, null, null);

            Assert.That(vm.Branches.Count, Is.EqualTo(2));
            Assert.That(vm.Branches[0].StateText, Is.EqualTo(L("Flow_Branch_Undecided")));
            Assert.That(vm.Branches[1].StateText, Is.EqualTo(L("Flow_Branch_Undecided")));
        }

        [Test]
        public void 一般流程_无去向且操作未登记_两分支均为待定()
        {
            FlowViewModel vm = FlowViewModel.BuildGeneral(Req(null), null, null, null);

            Assert.That(vm.Branches[0].StateText, Is.EqualTo(L("Flow_Branch_Undecided")));
            Assert.That(vm.Branches[1].StateText, Is.EqualTo(L("Flow_Branch_Undecided")));
        }
    }
}

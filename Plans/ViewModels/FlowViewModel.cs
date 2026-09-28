using ORT一键报告.Models;
using ORT一键报告.Services;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ORT一键报告.Plans.ViewModels
{
    /// <summary>
    /// 流程步骤状态常量（界面按此配色：完成=绿 / 当前=黄 / 未到达=灰 / 无需=灰空心 / 驳回=红）
    /// </summary>
    public static class FlowState
    {
        public const string Done = "done";
        public const string Current = "current";
        public const string Waiting = "waiting";
        public const string Skipped = "skipped";
        public const string Rejected = "rejected";
    }

    /// <summary>
    /// 流程里的一个步骤（节点）
    /// </summary>
    public class FlowStepItem
    {
        /// <summary>步骤名</summary>
        public string Title { get; set; }

        /// <summary>状态（见 <see cref="FlowState"/>）</summary>
        public string State { get; set; }

        /// <summary>状态词（已完成/进行中/未开始/无需/已驳回）</summary>
        public string StateText { get; set; }

        /// <summary>依据文本（日期/单号/扫描结果等；可为空）</summary>
        public string Evidence { get; set; }

        /// <summary>可点击链接文字（如「打開序列號文件」；为空表示没有链接）</summary>
        public string LinkText { get; set; }

        /// <summary>链接目标文件绝对路径</summary>
        public string LinkPath { get; set; }

        /// <summary>是否有可点击链接</summary>
        public bool HasLink => !string.IsNullOrWhiteSpace(LinkText) && !string.IsNullOrWhiteSpace(LinkPath);
    }

    /// <summary>
    /// 一步流程行：Items 为 1 个（普通节点）或 2 个（并行节点，如"新建報告模板 / 進行試驗"）
    /// </summary>
    public class FlowRow
    {
        public List<FlowStepItem> Items { get; set; } = [];
    }

    /// <summary>
    /// 流程分支（报废 / 回線入庫）
    /// </summary>
    public class FlowBranchItem
    {
        /// <summary>分支名</summary>
        public string Title { get; set; }

        /// <summary>分支描述（走此分支 / 未走此分支 / 待定）</summary>
        public string StateText { get; set; }

        /// <summary>是否走此分支（界面高亮）</summary>
        public bool IsActive { get; set; }

        /// <summary>是否置灰（未走此分支时）</summary>
        public bool IsDimmed { get; set; }

        /// <summary>分支内步骤</summary>
        public List<FlowStepItem> Steps { get; set; } = [];
    }

    /// <summary>
    /// 流程查看窗口的展示模型：步骤状态全部从现有登记数据推导（不新增状态字段），
    /// 由静态工厂 <see cref="BuildGeneral"/>（一般流程：领用/RT 计划）与
    /// <see cref="BuildExternalTest"/>（其他部门申请测试流程：QRT 计划）构造。
    /// </summary>
    public class FlowViewModel
    {
        /// <summary>流程类型名（一般流程（領用）/ 其他部門申請測試流程）</summary>
        public string FlowTitle { get; set; }

        /// <summary>记录标识（機種/單据號 或 工作編號/機種）</summary>
        public string Subtitle { get; set; }

        /// <summary>当前所处节点提示</summary>
        public string CurrentText { get; set; }

        /// <summary>分支前的线性步骤行</summary>
        public List<FlowRow> Rows { get; set; } = [];

        /// <summary>分支（无分支的流程为空）</summary>
        public List<FlowBranchItem> Branches { get; set; } = [];

        /// <summary>共用结案节点</summary>
        public FlowStepItem CloseStep { get; set; }

        /// <summary>底部判定依据说明</summary>
        public string FooterNote { get; set; }

        /* ###############################  工厂：一般流程  ################################ */

        /// <summary>
        /// 一般流程：單體領用+新增計劃 → 新建報告模板/進行試驗（可调换顺序）→ 完成報告 →
        /// （報廢）報廢前審核→報廢 或（入庫）回線→入庫 → 結案。
        /// </summary>
        /// <param name="req">领退记录（计划侧入口找不到领用记录时为 null）</param>
        /// <param name="plan">关联计划记录（领退侧入口找不到计划时为 null）</param>
        /// <param name="link">JobNo 对应的报告夹信息（未扫描到为 null）</param>
        /// <param name="scrapRequest">该领退记录最新的一条报废审核请求（没有为 null）</param>
        /// <param name="resolveAttachment">OleDir 相对路径 → 绝对路径（用于报废序列号文件链接）</param>
        public static FlowViewModel BuildGeneral(Requisition req, Plan plan, ReportLink link,
            ReviewRequest scrapRequest, Func<string, string> resolveAttachment = null)
        {
            FlowViewModel vm = new()
            {
                FlowTitle = L("Flow_TypeGeneral"),
                Subtitle = plan != null
                    ? F(L("Flow_SubtitlePlanFormat"), plan.JobNo ?? "-", plan.ModelName ?? "-")
                    : F(L("Flow_SubtitleReqFormat"), req?.ModelName ?? "-", req?.RequisitionNo ?? "-"),
                FooterNote = L("Flow_FooterNote")
            };

            // 1 單體領用
            vm.Rows.Add(Row(req != null
                ? Step("Flow_Step_Requisition", FlowState.Done,
                    F(L("Flow_Ev_RequisitionFormat"), req.RequisitionNo ?? "-", req.RequisitionDate))
                : Step("Flow_Step_Requisition", FlowState.Waiting, L("Flow_Ev_NoRequisition"))));

            // 2 新增計劃
            vm.Rows.Add(Row(plan != null
                ? Step("Flow_Step_Plan", FlowState.Done, F(L("Flow_Ev_PlanFormat"), plan.JobNo ?? "-"))
                : Step("Flow_Step_Plan", FlowState.Waiting, L("Flow_Ev_NoPlan"))));

            // 3+4 新建報告模板 / 進行試驗（并行，可调换顺序）
            FlowStepItem template = link != null
                ? Step("Flow_Step_Template", FlowState.Done,
                    F(L("Flow_Ev_TemplateFolderFormat"), string.IsNullOrWhiteSpace(link.ReportDir) ? "-" : Path.GetFileName(link.ReportDir)))
                : Step("Flow_Step_Template", FlowState.Waiting, L("Flow_Ev_TemplateMissing"));
            FlowStepItem test = DeriveTestStep(plan, "Flow_Step_Test");
            vm.Rows.Add(new FlowRow { Items = [template, test] });

            // 5 完成報告
            vm.Rows.Add(Row(DeriveReportStep(plan, "Flow_Step_Report")));

            // 分支 A：報廢
            FlowBranchItem scrapBranch = new() { Title = L("Flow_Branch_Scrap") };
            scrapBranch.Steps.Add(DeriveScrapReviewStep(req, scrapRequest));
            scrapBranch.Steps.Add(DeriveScrapStep(req, resolveAttachment));

            // 分支 B：回線 → 入庫
            FlowBranchItem stockInBranch = new() { Title = L("Flow_Branch_StockIn") };
            stockInBranch.Steps.Add(DeriveReturnStep(req));
            stockInBranch.Steps.Add(DeriveStockInStep(req));

            // 分支状态：报废字段已填=走报废；入库信息已填=走入库；都未填=待定；两者都有=异常提示
            bool scrapDone = req != null && (req.ScrapDate != null || !string.IsNullOrWhiteSpace(req.ScrapQty));
            bool stockInDone = req != null && (req.StockInDate != null || !string.IsNullOrWhiteSpace(req.StockInNo));
            ApplyBranchState(scrapBranch, scrapDone, stockInDone && !scrapDone);
            ApplyBranchState(stockInBranch, stockInDone, scrapDone && !stockInDone);
            if (scrapDone && stockInDone)
            {
                vm.FooterNote = L("Flow_WarnBothBranches") + Environment.NewLine + vm.FooterNote;
            }
            vm.Branches.Add(scrapBranch);
            vm.Branches.Add(stockInBranch);

            // 6 結案（共用节点）
            vm.CloseStep = DeriveCloseStep(plan);
            vm.CurrentText = BuildCurrentText(vm);
            return vm;
        }

        /* ###############################  工厂：其他部门申请测试流程  ################################ */

        /// <summary>
        /// 其他部门申请测试流程（QRT 计划）：登記信息 → 做測試 →（如需要）完成報告 → 結案
        /// </summary>
        public static FlowViewModel BuildExternalTest(Plan plan)
        {
            FlowViewModel vm = new()
            {
                FlowTitle = L("Flow_TypeExternal"),
                Subtitle = F(L("Flow_SubtitlePlanFormat"), plan?.JobNo ?? "-", plan?.ModelName ?? "-"),
                FooterNote = L("Flow_FooterNote")
            };

            // 1 登記信息
            string registerEvidence = plan == null
                ? L("Flow_Ev_NoPlan")
                : F(L("Flow_Ev_RegisterFormat"), plan.JobNo ?? "-", plan.ModelName ?? "-", plan.CreatedAt);
            vm.Rows.Add(Row(Step("Flow_Step_Register", plan != null ? FlowState.Done : FlowState.Waiting, registerEvidence)));

            // 2 做測試
            vm.Rows.Add(Row(DeriveTestStep(plan, "Flow_Step_TestExt")));

            // 3 完成報告（如需要）
            vm.Rows.Add(Row(DeriveReportStep(plan, "Flow_Step_ReportOptional")));

            // 4 結案
            vm.CloseStep = DeriveCloseStep(plan);
            vm.CurrentText = BuildCurrentText(vm);
            return vm;
        }

        /* ###############################  各节点推导  ################################ */

        /// <summary>
        /// 進行試驗 / 做測試：以报告扫描状态为主，兜底用计划状态与开始日期（依据文本里说明来源）
        /// </summary>
        private static FlowStepItem DeriveTestStep(Plan plan, string titleKey)
        {
            if (plan == null)
            {
                return Step(titleKey, FlowState.Waiting, L("Flow_Ev_NoPlan"));
            }
            string status = plan.ReportStatus;
            if (status == ReportStatusKind.Complete)
            {
                return Step(titleKey, FlowState.Done, L("Flow_Ev_TestDone"));
            }
            if (status == ReportStatusKind.InProgress)
            {
                return Step(titleKey, FlowState.Current, L("Flow_Ev_TestOngoing"));
            }
            if (status == ReportStatusKind.NotRequired)
            {
                return Step(titleKey, FallbackByPlanStatus(plan), L("Flow_Ev_TestNoReport"));
            }
            // 尚未扫描到报告数据：按计划状态/开始日期兜底
            string byStatus = FallbackByPlanStatus(plan);
            if (byStatus == FlowState.Waiting && plan.StartDate != null)
            {
                string evidence = plan.StartDate.Value.Date <= DateTime.Today
                    ? F(L("Flow_Ev_TestStartPassedFormat"), plan.StartDate)
                    : F(L("Flow_Ev_TestStartFutureFormat"), plan.StartDate);
                return Step(titleKey, byStatus, evidence);
            }
            string fallbackEvidence = byStatus == FlowState.Waiting ? L("Flow_Ev_TestNoStart")
                : F(L("Flow_Ev_PlanStatusFormat"), plan.Status ?? "-");
            return Step(titleKey, byStatus, fallbackEvidence);
        }

        /// <summary>计划状态兜底：已结案→完成 / 進行中→当前 / 预排→未到达 / 认不出→未到达</summary>
        private static string FallbackByPlanStatus(Plan plan)
            => plan == null ? FlowState.Waiting : PlanStatusKind.Of(plan.Status) switch
            {
                PlanStatusKind.Closed => FlowState.Done,
                PlanStatusKind.Ongoing => FlowState.Current,
                _ => FlowState.Waiting
            };

        /// <summary>
        /// 完成報告：已完成→完成；無要求→無需（跳过）；進行中→进行中；未扫描→未到达
        /// </summary>
        private static FlowStepItem DeriveReportStep(Plan plan, string titleKey)
        {
            if (plan == null)
            {
                return Step(titleKey, FlowState.Waiting, L("Flow_Ev_NoPlan"));
            }
            return plan.ReportStatus switch
            {
                ReportStatusKind.Complete => Step(titleKey, FlowState.Done, L("Flow_Ev_ReportComplete")),
                ReportStatusKind.InProgress => Step(titleKey, FlowState.Current, L("Flow_Ev_ReportOngoing")),
                ReportStatusKind.NotRequired => Step(titleKey, FlowState.Skipped, L("Flow_Ev_ReportNotRequired")),
                _ => Step(titleKey, FlowState.Waiting, L("Flow_Ev_ReportNotScanned"))
            };
        }

        /// <summary>
        /// 報廢前審核：已通过→完成（审核人/时间）；待审核→当前（审核人）；已驳回→驳回（意见）；
        /// 已有报废登记且无审核记录→完成（免审直登）；否则未到达
        /// </summary>
        private static FlowStepItem DeriveScrapReviewStep(Requisition req, ReviewRequest scrapRequest)
        {
            bool scrapRegistered = req != null && (req.ScrapDate != null || !string.IsNullOrWhiteSpace(req.ScrapQty));
            if (scrapRequest != null)
            {
                switch (scrapRequest.Status)
                {
                    case ReviewService.StatusApproved:
                        return Step("Flow_Step_ScrapReview", FlowState.Done,
                            F(L("Flow_Ev_ScrapReviewApprovedFormat"), scrapRequest.ReviewerName ?? "-", scrapRequest.ReviewedAt));
                    case ReviewService.StatusPending:
                        return Step("Flow_Step_ScrapReview", FlowState.Current,
                            F(L("Flow_Ev_ScrapReviewPendingFormat"), scrapRequest.AssigneeName ?? "-", scrapRequest.CreatedAt));
                    case ReviewService.StatusRejected:
                        return Step("Flow_Step_ScrapReview", FlowState.Rejected,
                            F(L("Flow_Ev_ScrapReviewRejectedFormat"), scrapRequest.ReviewComment ?? "-"));
                }
            }
            if (scrapRegistered)
            {
                return Step("Flow_Step_ScrapReview", FlowState.Done, L("Flow_Ev_ScrapReviewFree"));
            }
            return Step("Flow_Step_ScrapReview", FlowState.Waiting, L("Flow_Ev_ScrapReviewWaiting"));
        }

        /// <summary>
        /// 報廢：报废数量/日期已填→完成（数量/日期/单号/序列号个数；文件模式提供打开链接）；否则未到达
        /// </summary>
        private static FlowStepItem DeriveScrapStep(Requisition req, Func<string, string> resolveAttachment)
        {
            if (req == null || (req.ScrapDate == null && string.IsNullOrWhiteSpace(req.ScrapQty)))
            {
                return Step("Flow_Step_Scrap", FlowState.Waiting, null);
            }
            string evidence = F(L("Flow_Ev_ScrapFormat"), req.ScrapQty ?? "-", req.ScrapDate);
            if (!string.IsNullOrWhiteSpace(req.ScrapNo))
            {
                evidence += F(L("Flow_Ev_ScrapNoFormat"), req.ScrapNo);
            }
            FlowStepItem item = Step("Flow_Step_Scrap", FlowState.Done, evidence);
            if (!string.IsNullOrWhiteSpace(req.ScrapSnFilePath))
            {
                item.Evidence += F(L("Flow_Ev_ScrapSnFileFormat"), Path.GetFileName(req.ScrapSnFilePath));
                string fullPath = resolveAttachment?.Invoke(req.ScrapSnFilePath);
                if (!string.IsNullOrWhiteSpace(fullPath) && File.Exists(fullPath))
                {
                    item.LinkText = L("Flow_Link_OpenScrapSn");
                    item.LinkPath = fullPath;
                }
            }
            else if (!string.IsNullOrWhiteSpace(req.ScrapSnText))
            {
                int count = req.ScrapSnText.Count(c => c == '\n') + 1;
                item.Evidence += F(L("Flow_Ev_ScrapSnFormat"), count);
            }
            return item;
        }

        /// <summary>
        /// 回線：回线日期已填→完成；仅回线RT工令（需要回线未回）→当前；否则未到达
        /// </summary>
        private static FlowStepItem DeriveReturnStep(Requisition req)
        {
            if (req == null)
            {
                return Step("Flow_Step_ReturnLine", FlowState.Waiting, null);
            }
            if (req.ReturnDate != null)
            {
                return Step("Flow_Step_ReturnLine", FlowState.Done,
                    F(L("Flow_Ev_ReturnDoneFormat"), req.ReturnRtOrder ?? "-", req.ReturnDate));
            }
            if (!string.IsNullOrWhiteSpace(req.ReturnRtOrder))
            {
                return Step("Flow_Step_ReturnLine", FlowState.Current,
                    F(L("Flow_Ev_ReturnPendingFormat"), req.ReturnRtOrder));
            }
            return Step("Flow_Step_ReturnLine", FlowState.Waiting, null);
        }

        /// <summary>入庫：入库日期/单据号已填→完成（单据号/数量/日期）；否则未到达</summary>
        private static FlowStepItem DeriveStockInStep(Requisition req)
        {
            if (req == null || (req.StockInDate == null && string.IsNullOrWhiteSpace(req.StockInNo)))
            {
                return Step("Flow_Step_StockIn", FlowState.Waiting, null);
            }
            return Step("Flow_Step_StockIn", FlowState.Done,
                F(L("Flow_Ev_StockInFormat"), req.StockInNo ?? "-", req.StockInQty ?? "-", req.StockInDate));
        }

        /// <summary>結案：关联计划状态归类为已结案→完成；无关联计划→无法判定；否则未到达</summary>
        private static FlowStepItem DeriveCloseStep(Plan plan)
        {
            if (plan == null)
            {
                return Step("Flow_Step_Close", FlowState.Waiting, L("Flow_Ev_CloseNoPlan"));
            }
            return PlanStatusKind.Of(plan.Status) == PlanStatusKind.Closed
                ? Step("Flow_Step_Close", FlowState.Done, L("Flow_Ev_CloseDone"))
                : Step("Flow_Step_Close", FlowState.Waiting, L("Flow_Ev_CloseWaiting"));
        }

        /// <summary>
        /// 分支状态：本分支已完成→走此分支（高亮）；另一分支已完成→未走此分支（置灰）；都未完成为待定
        /// </summary>
        private static void ApplyBranchState(FlowBranchItem branch, bool selfDone, bool otherDone)
        {
            if (selfDone)
            {
                branch.StateText = L("Flow_Branch_Taken");
                branch.IsActive = true;
            }
            else if (otherDone)
            {
                branch.StateText = L("Flow_Branch_NotTaken");
                branch.IsDimmed = true;
            }
            else
            {
                branch.StateText = L("Flow_Branch_Undecided");
            }
        }

        private static string BuildCurrentText(FlowViewModel vm)
        {
            List<FlowStepItem> all = [];
            foreach (FlowRow row in vm.Rows)
            {
                all.AddRange(row.Items);
            }
            foreach (FlowBranchItem branch in vm.Branches.Where(b => !b.IsDimmed))
            {
                all.AddRange(branch.Steps);
            }
            if (vm.CloseStep != null)
            {
                all.Add(vm.CloseStep);
            }
            List<string> current = all.Where(s => s.State == FlowState.Current).Select(s => s.Title).ToList();
            if (current.Count > 0)
            {
                return F(L("Flow_CurrentFormat"), string.Join("、", current));
            }
            return all.All(s => s.State != FlowState.Waiting && s.State != FlowState.Rejected)
                ? L("Flow_CurrentAllDone")
                : L("Flow_CurrentNone");
        }

        /* ###############################  小工具  ################################ */

        private static string L(string key) => LanguageService.Get(key);

        private static string F(string format, params object[] args)
        {
            try
            {
                return string.Format(format, args);
            }
            catch
            {
                return format;
            }
        }

        private static FlowStepItem Step(string titleKey, string state, string evidence)
            => new() { Title = L(titleKey), State = state, StateText = StateTextOf(state), Evidence = evidence };

        private static string StateTextOf(string state) => state switch
        {
            FlowState.Done => L("Flow_State_Done"),
            FlowState.Current => L("Flow_State_Current"),
            FlowState.Waiting => L("Flow_State_Waiting"),
            FlowState.Skipped => L("Flow_State_Skipped"),
            FlowState.Rejected => L("Flow_State_Rejected"),
            _ => ""
        };

        private static FlowRow Row(FlowStepItem item) => new() { Items = [item] };
    }
}

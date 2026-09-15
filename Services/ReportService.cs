using ORT一键报告.Models;
using ORT一键报告.Reports.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 报告服务，封装原 WindowMainReport 的静态属性，
    /// 消除 ViewModel 对 View 层的静态依赖。
    /// </summary>
    public class ReportService
    {
        public string RootPath { get; set; }
        public string TemplateDir { get; set; }
        public string TempPath { get; set; }
        public UUTInfoFromExcel UUTInfos { get; set; }

        /// <summary>
        /// 根据报告文件夹名称（机种/RT工号等）从领退和计划中匹配到的记录，
        /// 用于补充报告表头信息（项目名/阶段/负责人等）
        /// </summary>
        public Plan MatchedPlan { get; set; }

        /// <summary>
        /// 匹配计划对应的领退记录（从计划表右键打开一键报告时携带），
        /// 用于补充 S/N 等领退数据到报告相应位置
        /// </summary>
        public Requisition MatchedRequisition { get; set; }

        /// <summary>
        /// 匹配计划绑定的报告文件夹（report_links.ReportDir，从计划表右键打开时携带）。
        /// 测试信息（TESTED BY/APPROVED BY…）与测试图片（Issue Photos/Test Setup）
        /// 优先从该文件夹下按报告类型的本地报告文件读取，读不到才回退模板。
        /// </summary>
        public string MatchedReportDir { get; set; }

        /// <summary>
        /// 本次是否"从计划表右键进入"（false = 直接进入一键报告）。
        /// 直接进入时不得用上一次残留的匹配记录覆盖报告概览读到的 SN/工令/版本。
        /// </summary>
        public bool EnteredFromPlan { get; set; }

        /// <summary>
        /// 清空跨窗口残留的匹配信息。ReportService 是单例，上一次从计划表进入留下的
        /// MatchedPlan/MatchedRequisition/MatchedReportDir/预填模型 会串到这一次，
        /// 因此"直接进入一键报告"时必须先清掉。
        /// </summary>
        public void ClearMatchedSource()
        {
            MatchedPlan = null;
            MatchedRequisition = null;
            MatchedReportDir = null;
            PrefilledReportModel = null;
            UUTInfos = null;
            EnteredFromPlan = false;
        }

        /// <summary>
        /// 预填的一键报告输入模型实例（从计划 + 领退构建）。
        /// 只要提供该实例即可直接生成报告，与 UI 解耦。
        /// </summary>
        public ReportInputModel PrefilledReportModel { get; set; }

        /// <summary>
        /// 用领用表（+计划表）的数据覆盖报告概览里读到的 UUT 信息：
        /// 从计划表右键进入时，序列号/工令/版本/DC 以领用表为准，"测试项目"等仍保留概览里的。
        /// 未从计划表进入（MatchedRequisition 为空）时不改动任何内容。
        /// </summary>
        /// <returns>是否发生了覆盖</returns>
        public bool ApplyMatchedSourceToUUTInfos()
        {
            UUTInfoFromExcel infos = UUTInfos;
            Requisition req = MatchedRequisition;
            // 只有"从计划表进入"时才用领用表覆盖；直接进入时报告概览才是数据来源
            if (infos == null || req == null || !EnteredFromPlan)
            {
                return false;
            }

            List<string> sns = (req.SN ?? "")
                .Split(['\n', '\r'], StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim())
                .Where(s => s.Length > 0)
                .ToList();

            bool changed = false;
            if (sns.Count > 0 && !sns.SequenceEqual(infos.SNs ?? []))
            {
                infos.SNs = sns;
                changed = true;
            }
            if (!string.IsNullOrWhiteSpace(req.WorkOrder) && req.WorkOrder != infos.WorkOrder)
            {
                infos.WorkOrder = req.WorkOrder;
                changed = true;
            }
            if (!string.IsNullOrWhiteSpace(req.Rev) && req.Rev != infos.Revision)
            {
                infos.Revision = req.Rev;
                changed = true;
            }
            if (!string.IsNullOrWhiteSpace(req.DC) && req.DC != infos.DC)
            {
                infos.DC = req.DC;
                changed = true;
            }
            return changed;
        }
    }
}
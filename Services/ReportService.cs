using ORT一键报告.Models;
using ORT一键报告.Reports.Models;

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
        /// 预填的一键报告输入模型实例（从计划 + 领退构建）。
        /// 只要提供该实例即可直接生成报告，与 UI 解耦。
        /// </summary>
        public ReportInputModel PrefilledReportModel { get; set; }
    }
}
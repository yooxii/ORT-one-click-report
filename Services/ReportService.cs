using ORT一键报告.Models;

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
    }
}
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
    }
}
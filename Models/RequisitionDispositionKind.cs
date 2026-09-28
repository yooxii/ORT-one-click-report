namespace ORT一键报告.Models
{
    /// <summary>
    /// 领退记录的单体去向（Requisition.Disposition）取值：整笔领用只走一条终态分支
    /// （报废 或 回线入库），在领退表新增/编辑时必选；流程查看据此在操作尚未登记时
    /// 标出计划走的分支。值直接落库为字符串，界面显示与存储一致（与 <see cref="ReportStatusKind"/> 同一模式）。
    /// </summary>
    public static class RequisitionDispositionKind
    {
        /// <summary>入库：测试后回线并入库（走回线入库分支）</summary>
        public const string StockIn = "入库";

        /// <summary>报废：测试后报废（走报废分支）</summary>
        public const string Scrap = "报废";

        /// <summary>下拉框选项（按"入库 → 报废"排序）</summary>
        public static readonly string[] All = [StockIn, Scrap];
    }
}

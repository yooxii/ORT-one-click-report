using ORT一键报告.Models;
using System;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// 表格字段 → 显示文本的统一取值：领退表/计划表的右键「复制单元格」、
    /// 批量窗口清单的右键「复制一次值」共用，避免两处各写一套 switch 出现偏差。
    /// 字段名用实体属性名（与 DataGridColumn.SortMemberPath 一致）。
    /// </summary>
    internal static class PlanFieldText
    {
        /// <summary>领退记录某字段的显示文本（日期按 yyyy/M/d 显示；SN 附件形式时给出文件名/路径）</summary>
        public static string RequisitionText(Requisition r, string field) => r == null ? null : field switch
        {
            nameof(Requisition.RequisitionDate) => r.RequisitionDate?.ToString("yyyy/M/d"),
            nameof(Requisition.RequisitionNo) => r.RequisitionNo,
            nameof(Requisition.ModelName) => r.ModelName,
            nameof(Requisition.OutQty) => r.OutQty,
            nameof(Requisition.Disposition) => r.Disposition,
            nameof(Requisition.SN) => r.SN ?? r.SnFilePath,
            nameof(Requisition.SnFilePath) => r.SnFilePath,
            nameof(Requisition.DC) => r.DC,
            nameof(Requisition.Rev) => r.Rev,
            nameof(Requisition.WorkOrder) => r.WorkOrder,
            nameof(Requisition.ReturnRtOrder) => r.ReturnRtOrder,
            nameof(Requisition.ReturnQty) => r.ReturnQty,
            nameof(Requisition.LineNo) => r.LineNo,
            nameof(Requisition.ReturnDate) => r.ReturnDate?.ToString("yyyy/M/d"),
            nameof(Requisition.StockInNo) => r.StockInNo,
            nameof(Requisition.StockInQty) => r.StockInQty,
            nameof(Requisition.StockInDate) => r.StockInDate?.ToString("yyyy/M/d"),
            nameof(Requisition.ScrapNo) => r.ScrapNo,
            nameof(Requisition.ScrapQty) => r.ScrapQty,
            nameof(Requisition.ScrapDate) => r.ScrapDate?.ToString("yyyy/M/d"),
            nameof(Requisition.Remark) => r.Remark,
            _ => null
        };

        /// <summary>计划记录某字段的显示文本（日期按 yyyy/M/d 显示）</summary>
        public static string PlanText(Plan p, string field) => p == null ? null : field switch
        {
            nameof(Plan.JobNo) => p.JobNo,
            nameof(Plan.Product) => p.Product,
            nameof(Plan.Customer) => p.Customer,
            nameof(Plan.ModelName) => p.ModelName,
            nameof(Plan.Stage) => p.Stage,
            nameof(Plan.TestItem) => p.TestItem,
            nameof(Plan.SampleSize) => p.SampleSize,
            nameof(Plan.TestPeriod) => p.TestPeriod,
            nameof(Plan.Owner) => p.Owner,
            nameof(Plan.StartDate) => p.StartDate?.ToString("yyyy/M/d"),
            nameof(Plan.EndDate) => p.EndDate?.ToString("yyyy/M/d"),
            nameof(Plan.Status) => p.Status,
            nameof(Plan.ReportStatus) => p.ReportStatus,
            nameof(Plan.UnitReturnDate) => p.UnitReturnDate?.ToString("yyyy/M/d"),
            nameof(Plan.Remark) => p.Remark,
            _ => null
        };
    }
}

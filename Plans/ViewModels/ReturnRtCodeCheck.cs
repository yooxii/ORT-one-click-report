using ORT一键报告.Models;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ORT一键报告.Plans.ViewModels
{
    /// <summary>
    /// 领退表「回线RT工令」编号体检（只读，不改数据）：
    /// 列出编号月段与领用日期所在年月不一致的记录（多为历史数据或导入时写错，
    /// 会让该年月的新编号从错误的序号往后接），以及根本不是编号形状的值。
    /// 判定规则见 <see cref="RequisitionEditRules.InspectReturnRt"/>。
    /// </summary>
    public static class ReturnRtCodeCheck
    {
        /// <summary>
        /// 扫一遍领退表，返回（检查了多少条、有问题的记录）。
        /// 问题记录按领用日期、Id 排序，便于与源表逐条核对。
        /// </summary>
        public static (int Scanned, List<ReturnRtCodeIssue> Issues) Find(DatabaseService db)
        {
            List<Requisition> requisitions = db.FreeSql.Select<Requisition>()
                .Where(r => r.ReturnRtOrder != null && r.ReturnRtOrder != "")
                .ToList();
            List<ReturnRtCodeIssue> issues = [];
            foreach (Requisition requisition in requisitions)
            {
                (ReturnRtCodeIssueKind Kind, string SuggestedCode)? found = RequisitionEditRules.InspectReturnRt(
                    requisition.ReturnRtOrder, requisition.RequisitionDate, requisition.ReturnDate);
                if (found == null)
                {
                    continue;
                }
                issues.Add(new ReturnRtCodeIssue
                {
                    Id = requisition.Id,
                    RequisitionDate = requisition.RequisitionDate,
                    ReturnDate = requisition.ReturnDate,
                    RequisitionNo = requisition.RequisitionNo,
                    ModelName = requisition.ModelName,
                    ReturnRtOrder = requisition.ReturnRtOrder,
                    Kind = found.Value.Kind,
                    CodeYearMonth = RequisitionEditRules.ReturnRtYearMonth(requisition.ReturnRtOrder),
                    SuggestedCode = found.Value.SuggestedCode
                });
            }
            return (requisitions.Count, issues.OrderBy(i => i.RequisitionDate ?? DateTime.MaxValue)
                .ThenBy(i => i.Id)
                .ToList());
        }
    }
}

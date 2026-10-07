using ORT一键报告.Models;
using ORT一键报告.Services;

namespace ORT一键报告.Plans.ViewModels
{
    /// <summary>
    /// 领退表批量操作类型（多选右键菜单用）
    /// </summary>
    public enum RequisitionBatchMode
    {
        /// <summary>批量回线：把同一个回线日期写入多条领退记录</summary>
        Return,

        /// <summary>批量入库：把同一组入库单据/数量/日期写入多条领退记录</summary>
        StockIn,

        /// <summary>批量标记删除：把多条领退记录标记删除（暂存，「提交保存」后生效）</summary>
        Delete
    }

    /// <summary>
    /// 领退表批量登记的可执行判定：与单条右键动作的分支限制完全一致
    /// （报废去向不能回线/入库、报废后不能入库、未回线不能入库），
    /// 批量窗口据此逐条给出「可执行 / 跳过原因」，只对可执行的记录写入。
    /// </summary>
    public static class RequisitionBatchRules
    {
        /// <summary>
        /// 批量回线：不可执行时返回原因文本，可执行返回 null
        /// </summary>
        public static string ReturnBlockReason(Requisition req)
        {
            if (req == null)
            {
                return LanguageService.Get("Plans_Msg_SelectRequisition");
            }
            // 单体去向与操作分支一致：报废的记录不能回线（回线属于入库分支）
            return req.Disposition != RequisitionDispositionKind.StockIn
                ? LanguageService.Get("Msg_OpBlockedDispositionScrap")
                : null;
        }

        /// <summary>
        /// 批量入库：不可执行时返回原因文本（报废去向 / 已报废 / 未回线），可执行返回 null
        /// </summary>
        public static string StockInBlockReason(Requisition req)
        {
            if (req == null)
            {
                return LanguageService.Get("Plans_Msg_SelectRequisition");
            }
            if (req.Disposition != RequisitionDispositionKind.StockIn)
            {
                return LanguageService.Get("Msg_OpBlockedDispositionScrap");
            }
            if (req.ScrapDate != null || !string.IsNullOrWhiteSpace(req.ScrapQty))
            {
                return LanguageService.Get("Msg_StockInAfterScrap");
            }
            return req.ReturnDate == null ? LanguageService.Get("Msg_StockInNeedReturn") : null;
        }

        /// <summary>
        /// 按批量操作类型取不可执行原因（删除没有分支限制，恒返回 null）
        /// </summary>
        public static string BlockReason(RequisitionBatchMode mode, Requisition req) => mode switch
        {
            RequisitionBatchMode.Return => ReturnBlockReason(req),
            RequisitionBatchMode.StockIn => StockInBlockReason(req),
            _ => null
        };
    }
}

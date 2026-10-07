using ORT一键报告.Models;
using ORT一键报告.Services;

namespace ORT一键报告.Plans.ViewModels
{
    /// <summary>
    /// 领退表批量操作类型（工具菜单的批量登记用）
    /// </summary>
    public enum RequisitionBatchMode
    {
        /// <summary>批量回线：把同一个回线日期写入多条领退记录</summary>
        Return,

        /// <summary>批量标记删除：把多条领退记录标记删除（暂存，「提交保存」后生效）</summary>
        Delete
    }

    /// <summary>
    /// 领退表登记规则的统一判定：与单条右键动作的分支限制完全一致
    /// （报废去向不能回线/入库、报废后不能入库、未回线不能入库）。
    /// 批量窗口据此逐条给出「可执行 / 跳过原因」，工具菜单与逐条登记据此筛出待办记录。
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
        /// 入库：不可执行时返回原因文本（报废去向 / 已报废 / 未回线），可执行返回 null
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
            if (IsScrapped(req))
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
            _ => null
        };

        /* ###############################  表格筛选口径  ################################ */

        /// <summary>
        /// 待回线（批量回线时表格要筛出来的记录）：单体去向为入库，且还没登记回线日期
        /// </summary>
        public static bool NeedsReturn(Requisition req)
            => req != null && req.Disposition == RequisitionDispositionKind.StockIn && req.ReturnDate == null;

        /// <summary>
        /// 待入库（工具菜单「筛选待入库」筛出来的记录）：单体去向为入库、未报废、还没登记入库。
        /// 包含还没回线的记录 —— 它们同样属于「还要走入库这一步」，只是要先回线才能登记。
        /// </summary>
        public static bool NeedsStockIn(Requisition req)
            => req != null
               && req.Disposition == RequisitionDispositionKind.StockIn
               && !IsScrapped(req)
               && LacksStockIn(req);

        /// <summary>
        /// 可以马上逐条登记入库（入库窗口里「下一个待入库」只跳到这类记录）：
        /// 待入库且已经回线 —— 报废去向/已报废/未回线都登记不了，与单条右键「入库」的限制一致
        /// </summary>
        public static bool CanRegisterStockIn(Requisition req)
            => StockInBlockReason(req) == null && LacksStockIn(req);

        /// <summary>
        /// 批量登记时表格的临时筛选口径（删除没有「待办」口径，恒为 true，即不筛）
        /// </summary>
        public static bool MatchesQuickFilter(RequisitionBatchMode mode, Requisition req) => mode switch
        {
            RequisitionBatchMode.Return => NeedsReturn(req),
            _ => true
        };

        /* ###############################  口径复用的小判定  ################################ */

        /// <summary>已经登记过报废（报废日期或报废数量有值）</summary>
        private static bool IsScrapped(Requisition req)
            => req.ScrapDate != null || !string.IsNullOrWhiteSpace(req.ScrapQty);

        /// <summary>还没登记入库（入库单号与入库日期都是空的）</summary>
        private static bool LacksStockIn(Requisition req)
            => req.StockInDate == null && string.IsNullOrWhiteSpace(req.StockInNo);
    }
}

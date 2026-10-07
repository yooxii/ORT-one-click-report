using CommunityToolkit.Mvvm.ComponentModel;
using ORT一键报告.Models;
using ORT一键报告.Plans.ViewModels;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// 批量登记窗口清单的一行：包住领退记录，附带「本次结果」。
    /// 可执行性由 <see cref="RequisitionBatchRules"/> 判定并在打开窗口时固定；
    /// <see cref="Status"/> 随用户改动的日期/单据实时刷新（如「将写入回线日期 2026/10/7」）。
    /// </summary>
    public class RequisitionBatchRow : ObservableObject
    {
        public RequisitionBatchRow(RequisitionBatchMode mode, Requisition item)
        {
            Item = item;
            BlockReason = RequisitionBatchRules.BlockReason(mode, item);
        }

        /// <summary>被操作的领退记录</summary>
        public Requisition Item { get; }

        /// <summary>不可执行的原因；可执行时为 null</summary>
        public string BlockReason { get; }

        /// <summary>本次是否会写入（false 表示跳过）</summary>
        public bool Eligible => BlockReason == null;

        private string _status = "";
        /// <summary>「本次结果」列文本：将写入什么，或跳过原因</summary>
        public string Status { get => _status; set => SetProperty(ref _status, value); }
    }
}

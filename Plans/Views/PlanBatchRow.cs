using ORT一键报告.Models;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// 计划表批量删除窗口清单的一行：包住计划记录，附带「本次结果」文本
    /// （计划表本身已有「状况」字段，故结果列单独取属性名，避免与状况混淆）。
    /// </summary>
    public class PlanBatchRow
    {
        public PlanBatchRow(Plan item, string result)
        {
            Item = item;
            Result = result;
        }

        /// <summary>被操作的计划记录</summary>
        public Plan Item { get; }

        /// <summary>「本次结果」列文本</summary>
        public string Result { get; }
    }
}

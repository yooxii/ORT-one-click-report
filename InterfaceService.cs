namespace ORT一键报告
{
    public interface IEMIService
    {
        /// <summary>
        /// 打开EMI数据文件选择框
        /// </summary>
        /// <param name="title">对话框标题</param>
        /// <param name="filter">过滤选项</param>
        /// <param name="initPath">选择框打开的初始位置</param>
        /// <param name="isDir">true 返回文件夹路径; false 返回文件路径</param>
        /// <returns>返回选择的EMI数据路径；取消返回 null</returns>
        string OpenPathDialog(string title, string filter = "EMI模板|*.xlsx;*.xlsm|所有文件|*.*", string initPath = null, bool isDir = false);
    }
}
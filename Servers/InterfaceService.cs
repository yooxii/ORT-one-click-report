namespace ORT一键报告
{
    public interface IService
    {
        /// <summary>
        /// 打开文件选择框
        /// </summary>
        /// <param name="title">对话框标题</param>
        /// <param name="filter">过滤选项</param>
        /// <param name="initPath">选择框打开的初始位置</param>
        /// <param name="isDir">true 返回文件夹路径; false 返回文件路径</param>
        /// <returns>返回选择的路径；取消返回 null</returns>
        string OpenPathDialog(string title, string filter = "Excel文件|*.xls;*.xlsx;*.xlsm|所有文件|*.*", string initPath = null, bool isDir = false);

        /// <summary>
        /// 打开保存对话框
        /// </summary>
        /// <param name="title">对话框标题</param>
        /// <param name="saveName">保存的文件名</param>
        /// <param name="filter">过滤选项</param>
        /// <param name="initPath">选择框打开的初始位置</param>
        /// <returns>返回选择的保存路径；取消返回 null</returns>
        public string SavePathDialog(string title, string saveName = null, string filter = "Excel文件|*.xlsx|所有文件|*.*", string initPath = null);
    }
}
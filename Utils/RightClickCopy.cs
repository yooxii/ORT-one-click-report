using ORT一键报告.Services;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace ORT一键报告.Utils
{
    /// <summary>
    /// 「右键单击复制一次值」：右键落在只读文本/表格单元格上时，直接把该值复制到剪贴板
    /// （一次一个值），不再弹默认编辑菜单。用于各登记窗口的只读字段与批量窗口的清单。
    /// 输入框（非只读）保持系统默认行为，不影响粘贴等操作。
    /// </summary>
    public static class RightClickCopy
    {
        /// <summary>提示里最多显示的原值长度</summary>
        private const int PreviewLength = 60;

        /// <summary>
        /// 让窗口内所有只读 TextBox 支持右键单击即复制
        /// </summary>
        public static void AttachReadOnlyText(Window window)
        {
            if (window == null)
            {
                return;
            }
            window.PreviewMouseRightButtonDown += (s, e) =>
            {
                if (e.OriginalSource is not DependencyObject source)
                {
                    return;
                }
                TextBox box = FindAncestor<TextBox>(source);
                if (box == null || !box.IsReadOnly)
                {
                    return;
                }
                // 去掉系统编辑菜单：右键即复制，不再给第二个入口
                box.ContextMenu = null;
                if (Copy(box.Text))
                {
                    e.Handled = true;
                }
            };
        }

        /// <summary>
        /// 让表格支持右键单击单元格即复制该值；取值由 <paramref name="valueOf"/> 提供
        /// （参数为单元格绑定对象与该列，返回该单元格的完整值）
        /// </summary>
        public static void AttachDataGrid(DataGrid grid, Func<object, DataGridColumn, string> valueOf)
        {
            if (grid == null || valueOf == null)
            {
                return;
            }
            grid.PreviewMouseRightButtonDown += (s, e) =>
            {
                if (e.OriginalSource is not DependencyObject source)
                {
                    return;
                }
                DataGridCell cell = FindAncestor<DataGridCell>(source);
                if (cell?.Column == null)
                {
                    return;
                }
                if (Copy(valueOf(cell.DataContext, cell.Column)))
                {
                    e.Handled = true;
                }
            };
        }

        /// <summary>
        /// 复制文本并给出「已复制：xxx」提示；文本为空或复制失败返回 false
        /// </summary>
        public static bool Copy(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }
            try
            {
                Clipboard.SetText(text);
            }
            catch (Exception ex)
            {
                // 剪贴板被其他进程占用等：不打断用户操作
                NLog.LogManager.GetCurrentClassLogger().Warn($"复制到剪贴板失败: {ex.Message}");
                return false;
            }
            string preview = text.Length > PreviewLength ? text.Substring(0, PreviewLength) + "…" : text;
            ToastService.Show(string.Format(LanguageService.Get("Msg_ValueCopiedFormat"), preview));
            return true;
        }

        /// <summary>沿可视树向上查找指定类型的祖先</summary>
        private static T FindAncestor<T>(DependencyObject node) where T : DependencyObject
        {
            while (node != null)
            {
                if (node is T match)
                {
                    return match;
                }
                node = node is Visual || node is System.Windows.Media.Media3D.Visual3D
                    ? VisualTreeHelper.GetParent(node)
                    : LogicalTreeHelper.GetParent(node);
            }
            return null;
        }
    }
}

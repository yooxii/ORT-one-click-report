using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace ORT一键报告.Utils
{
    /// <summary>
    /// DataGrid 行拖动排序（计划明细、报告模板测试项都用它）。
    /// 用法：<c>DataGridRowDrag.Enable(grid, (from, to) =&gt; 交换/重排);</c>
    /// 拖动过程中按住 Ctrl 无效；只处理行本身，表头/滚动条不受影响。
    /// </summary>
    public static class DataGridRowDrag
    {
        /// <summary>
        /// 打开拖动排序：把 from 行拖到 to 行时回调（两个参数都是行数据对象）
        /// </summary>
        public static void Enable(DataGrid grid, Action<object, object> onMove, Func<object, bool> canDrag = null)
        {
            if (grid == null || onMove == null)
            {
                return;
            }
            Point start = default;
            DataGridRow dragRow = null;

            grid.AllowDrop = true;
            grid.PreviewMouseLeftButtonDown += (s, e) =>
            {
                start = e.GetPosition(null);
                dragRow = FindRow(e.OriginalSource as DependencyObject);
                if (dragRow != null && canDrag != null && !canDrag(dragRow.Item))
                {
                    dragRow = null;
                }
            };
            grid.PreviewMouseMove += (s, e) =>
            {
                if (dragRow == null || e.LeftButton != MouseButtonState.Pressed)
                {
                    return;
                }
                Point now = e.GetPosition(null);
                if (Math.Abs(now.X - start.X) < SystemParameters.MinimumHorizontalDragDistance
                    && Math.Abs(now.Y - start.Y) < SystemParameters.MinimumVerticalDragDistance)
                {
                    return;
                }
                object item = dragRow.Item;
                dragRow = null;
                try
                {
                    DragDrop.DoDragDrop(grid, item, DragDropEffects.Move);
                }
                catch
                {
                    // 拖动被系统取消（例如窗口失活）时忽略
                }
            };
            grid.DragOver += (s, e) =>
            {
                e.Effects = e.Data.GetFormats().Length > 0 ? DragDropEffects.Move : DragDropEffects.None;
                e.Handled = true;
            };
            grid.Drop += (s, e) =>
            {
                try
                {
                    object source = null;
                    foreach (string format in e.Data.GetFormats())
                    {
                        object data = e.Data.GetData(format);
                        if (data != null)
                        {
                            source = data;
                            break;
                        }
                    }
                    if (source == null)
                    {
                        return;
                    }
                    DataGridRow targetRow = FindRow(e.OriginalSource as DependencyObject);
                    if (targetRow == null || ReferenceEquals(targetRow.Item, source))
                    {
                        return;
                    }
                    onMove(source, targetRow.Item);
                }
                catch
                {
                    // 目标行不在表格内等情况直接忽略
                }
            };
        }

        /// <summary>
        /// 找鼠标位置所在的 DataGridRow（不在行上返回 null）
        /// </summary>
        private static DataGridRow FindRow(DependencyObject source)
        {
            while (source != null && source is not DataGridRow)
            {
                source = source is Visual or System.Windows.Media.Media3D.Visual3D
                    ? VisualTreeHelper.GetParent(source)
                    : LogicalTreeHelper.GetParent(source);
            }
            return source as DataGridRow;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 界面缩放：字号变大时，把「按像素写死」的布局一起按同一比例放大，
    /// 效果等同于 Windows 显示缩放（DPI 缩放）：窗口尺寸、表格列宽/行高、Grid 的绝对行列。
    /// 文字本身仍由 Window.FontSize 控制，不在这里二次放大。
    /// 说明：缩放比例 = 当前字号 / 基准字号（14）；窗口与表格各自记录已应用的比例，
    /// 因此重复调用只按增量缩放，不会累乘；后创建的窗口/表格会从自身设计尺寸一次性缩放到当前比例。
    /// 仅遍历逻辑树，模板内部（DataGrid/ComboBox 等自带模板）的尺寸不受影响。
    /// </summary>
    public static class UiScale
    {
        /// <summary>基准字号：界面的绝对像素尺寸按此字号设计</summary>
        public const double BaseFontSize = 14.0;

        /// <summary>最小缩放比例（字号调小时不无限缩小）</summary>
        public const double MinScale = 0.8;

        /// <summary>已应用比例（附加属性）：窗口/表格上记录当前生效的缩放值</summary>
        private static readonly DependencyProperty AppliedScaleProperty =
            DependencyProperty.RegisterAttached("AppliedScale", typeof(double), typeof(UiScale), new PropertyMetadata(0.0));

        /// <summary>本进程内最近一次生效的比例（供后创建的表格在 Loaded 时对齐）</summary>
        private static double _currentScale = 1.0;

        /// <summary>DataGrid.Loaded 类处理器是否已注册</summary>
        private static bool _gridLoadedHooked;

        /// <summary>
        /// 字号 → 缩放比例（钳制在 MinScale 与 maxScale 之间）
        /// </summary>
        public static double ScaleFor(double fontSize, double maxScale)
        {
            double scale = fontSize / BaseFontSize;
            if (scale < MinScale)
            {
                scale = MinScale;
            }
            if (scale > maxScale)
            {
                scale = maxScale;
            }
            return scale;
        }

        /// <summary>
        /// 把缩放应用到窗口：窗口尺寸、窗口内所有 DataGrid、所有 Grid 的绝对行列
        /// </summary>
        public static void Apply(Window window, double scale)
        {
            if (window == null || scale <= 0)
            {
                return;
            }
            _currentScale = scale;
            HookGridLoaded();
            double ratio = RatioFor(window, scale);
            if (ratio == 1.0)
            {
                return;
            }
            ScaleWindow(window, ratio);
            foreach (DataGrid grid in FindDescendants<DataGrid>(window))
            {
                ScaleGridColumns(grid, ratio);
                grid.SetValue(AppliedScaleProperty, scale);
            }
            foreach (Grid grid in FindDescendants<Grid>(window))
            {
                ScaleDefinitions(grid, ratio);
            }
            window.SetValue(AppliedScaleProperty, scale);
        }

        /* ###############################  内部实现  ################################ */

        /// <summary>
        /// 计算增量比例：目标比例 / 该对象已应用比例（未记录过视为 1）
        /// </summary>
        private static double RatioFor(DependencyObject target, double scale)
        {
            double applied = (double)target.GetValue(AppliedScaleProperty);
            if (applied <= 0)
            {
                applied = 1.0;
            }
            double ratio = scale / applied;
            return Math.Abs(ratio - 1.0) < 0.001 ? 1.0 : ratio;
        }

        /// <summary>
        /// 缩放窗口自身尺寸（用户手动调整过的大小也按增量缩放；最大化/自适应内容时不动）；
        /// 上限取当前监视器的工作区，避免放大后窗口跑到屏幕外（保存按钮不可达）
        /// </summary>
        private static void ScaleWindow(Window window, double ratio)
        {
            if (window.WindowState != WindowState.Normal || window.SizeToContent != SizeToContent.Manual)
            {
                return;
            }
            Rect work = AvailableUnits(window);
            if (!double.IsNaN(window.Width) && window.Width > 0)
            {
                window.Width = Clamp(window.Width * ratio, window.MinWidth, work.Width);
            }
            if (!double.IsNaN(window.Height) && window.Height > 0)
            {
                window.Height = Clamp(window.Height * ratio, window.MinHeight, work.Height);
            }
            // 放大后把窗口拉回可用工作区内（否则居中于所有者会顶到任务栏下方，底部按钮被遮住）
            if (!double.IsNaN(window.Left) && !double.IsNaN(window.Top))
            {
                window.Left = Math.Min(Math.Max(window.Left, work.Left), Math.Max(work.Left, work.Right - window.Width));
                window.Top = Math.Min(Math.Max(window.Top, work.Top), Math.Max(work.Top, work.Bottom - window.Height));
            }
        }

        /// <summary>
        /// 当前监视器的可用工作区，换算成窗口所用的逻辑单位。
        /// 逻辑单位与物理像素的比例由「窗口物理尺寸 / 窗口逻辑尺寸」实测得到，
        /// 因此不受进程 DPI 感知方式影响。
        /// </summary>
        private static Rect AvailableUnits(Window window)
        {
            Rect unlimited = new(0, 0, double.PositiveInfinity, double.PositiveInfinity);
            try
            {
                IntPtr handle = new WindowInteropHelper(window).Handle;
                if (handle == IntPtr.Zero)
                {
                    return unlimited;
                }
                MonitorInfo info = new() { Size = Marshal.SizeOf<MonitorInfo>() };
                if (!GetMonitorInfo(MonitorFromWindow(handle, MonitorDefaultToNearest), ref info)
                    || !GetWindowRect(handle, out NativeRect rect))
                {
                    return unlimited;
                }
                double physicalWidth = rect.Right - rect.Left;
                double physicalHeight = rect.Bottom - rect.Top;
                if (physicalWidth <= 0 || physicalHeight <= 0 || window.Width <= 0 || window.Height <= 0)
                {
                    return unlimited;
                }
                double scaleX = physicalWidth / window.Width;
                double scaleY = physicalHeight / window.Height;
                // 连同工作区原点一起换算，保证多显示器下窗口仍留在同一块屏幕上
                return new Rect(
                    info.Work.Left / scaleX,
                    info.Work.Top / scaleY,
                    (info.Work.Right - info.Work.Left) / scaleX,
                    (info.Work.Bottom - info.Work.Top) / scaleY);
            }
            catch
            {
                return unlimited;
            }
        }

        private static double Clamp(double value, double min, double max)
        {
            if (max > 0 && value > max)
            {
                value = max;
            }
            if (value < min)
            {
                value = min;
            }
            return value;
        }

        /// <summary>
        /// 缩放表格的绝对列宽/行高（Auto、星号列宽不受影响）
        /// </summary>
        private static void ScaleGridColumns(DataGrid grid, double ratio)
        {
            foreach (DataGridColumn column in grid.Columns)
            {
                if (column.Width.IsAbsolute && column.Width.Value > 0)
                {
                    column.Width = new DataGridLength(column.Width.Value * ratio);
                }
                if (column.MinWidth > 0)
                {
                    column.MinWidth *= ratio;
                }
                if (!double.IsInfinity(column.MaxWidth) && column.MaxWidth > 0)
                {
                    column.MaxWidth *= ratio;
                }
            }
            if (!double.IsNaN(grid.RowHeight) && grid.RowHeight > 0)
            {
                grid.RowHeight *= ratio;
            }
            if (!double.IsNaN(grid.ColumnHeaderHeight) && grid.ColumnHeaderHeight > 0)
            {
                grid.ColumnHeaderHeight *= ratio;
            }
        }

        /// <summary>
        /// 缩放 Grid 中以像素写死的行高/列宽，以及据此写死的 RowDefinition Height="20" 这类布局
        /// </summary>
        private static void ScaleDefinitions(Grid grid, double ratio)
        {
            foreach (RowDefinition row in grid.RowDefinitions)
            {
                if (row.Height.IsAbsolute && row.Height.Value > 0)
                {
                    row.Height = new GridLength(row.Height.Value * ratio);
                }
                if (row.MinHeight > 0)
                {
                    row.MinHeight *= ratio;
                }
                if (!double.IsInfinity(row.MaxHeight) && row.MaxHeight > 0)
                {
                    row.MaxHeight *= ratio;
                }
            }
            foreach (ColumnDefinition column in grid.ColumnDefinitions)
            {
                if (column.Width.IsAbsolute && column.Width.Value > 0)
                {
                    column.Width = new GridLength(column.Width.Value * ratio);
                }
                if (column.MinWidth > 0)
                {
                    column.MinWidth *= ratio;
                }
                if (!double.IsInfinity(column.MaxWidth) && column.MaxWidth > 0)
                {
                    column.MaxWidth *= ratio;
                }
            }
        }

        /// <summary>
        /// 沿逻辑树收集指定类型的后代（模板生成的可视子元素不在逻辑树中，避免误改控件模板内部尺寸）
        /// </summary>
        private static IEnumerable<T> FindDescendants<T>(DependencyObject root) where T : DependencyObject
        {
            if (root == null)
            {
                yield break;
            }
            foreach (object child in LogicalTreeHelper.GetChildren(root))
            {
                if (child is not DependencyObject node)
                {
                    continue;
                }
                if (node is T match)
                {
                    yield return match;
                }
                foreach (T nested in FindDescendants<T>(node))
                {
                    yield return nested;
                }
            }
        }

        /// <summary>
        /// 注册 DataGrid.Loaded 类处理器：切换 Tab、打开对话框等场景下延迟创建的表格，
        /// 在加载时补上当前缩放（否则会保持设计尺寸，字大时列宽不够）
        /// </summary>
        private static void HookGridLoaded()
        {
            if (_gridLoadedHooked)
            {
                return;
            }
            _gridLoadedHooked = true;
            EventManager.RegisterClassHandler(typeof(DataGrid), FrameworkElement.LoadedEvent,
                new RoutedEventHandler((sender, _) =>
                {
                    if (sender is not DataGrid grid)
                    {
                        return;
                    }
                    double ratio = RatioFor(grid, _currentScale);
                    if (ratio == 1.0)
                    {
                        return;
                    }
                    ScaleGridColumns(grid, ratio);
                    grid.SetValue(AppliedScaleProperty, _currentScale);
                }));
        }

        /* ###############################  原生互操作（取监示器工作区）  ################################ */

        /// <summary>MONITOR_DEFAULTTONEAREST</summary>
        private const uint MonitorDefaultToNearest = 2;

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeRect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MonitorInfo
        {
            public int Size;
            public NativeRect Monitor;
            public NativeRect Work;
            public uint Flags;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromWindow(IntPtr handle, uint flags);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool GetMonitorInfo(IntPtr monitor, ref MonitorInfo info);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr handle, out NativeRect rect);
    }
}

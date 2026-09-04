using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 窗口标题栏主题辅助类。
    /// 深色主题（DarkLab）下通过 DWM 沉浸式深色模式 API 将系统标题栏与窗口边框
    /// 统一切换为深色，与深色 UI 协调；浅色主题下还原系统默认浅色标题栏。
    /// 需要 Windows 10 2004 (build 19041) 及以上；更早版本静默跳过（标题栏保持系统默认）。
    /// </summary>
    public static class WindowThemeHelper
    {
        /// <summary>DWMWA_USE_IMMERSIVE_DARK_MODE（新值，Win10 20H1+）</summary>
        private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

        /// <summary>DWMWA_USE_IMMERSIVE_DARK_MODE（旧值，Win10 1809~1903）</summary>
        private const int DWMWA_USE_IMMERSIVE_DARK_MODE_OLD = 19;

        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_FRAMECHANGED = 0x0020;

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
            int x, int y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left, Top, Right, Bottom;
        }

        /// <summary>
        /// 按当前主题应用/还原窗口标题栏深色模式
        /// </summary>
        public static void ApplyToWindow(Window window)
        {
            try
            {
                IntPtr hwnd = new WindowInteropHelper(window).Handle;
                if (hwnd == IntPtr.Zero)
                {
                    return;
                }

                // 仅深色主题启用深色标题栏，其余主题还原浅色
                int useDark = ThemeService.CurrentTheme == "DarkLab" ? 1 : 0;

                // 先尝试新属性值(20)，失败再尝试旧值(19)，兼容不同 Win10 版本
                if (DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ref useDark, sizeof(int)) != 0)
                {
                    DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE_OLD, ref useDark, sizeof(int));
                }

                // 强制 DWM 重绘非客户区（标题栏/边框）
                SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0,
                    SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);

                // Win10 活动标题栏缓存仅靠 FRAMECHANGED 不会刷新（浅色→深色时前台窗口不生效）；
                // 对已显示窗口做 1px 宽度抖动，强制标题栏重排重画（首开窗口走 SourceInitialized 无需抖动）
                if (window.IsVisible && GetWindowRect(hwnd, out RECT r))
                {
                    int w = r.Right - r.Left;
                    int h = r.Bottom - r.Top;
                    SetWindowPos(hwnd, IntPtr.Zero, 0, 0, w + 1, h, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
                    SetWindowPos(hwnd, IntPtr.Zero, 0, 0, w, h, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
                }
            }
            catch
            {
                // 不支持 DWM 深色模式的系统（Win10 2004 以下）静默跳过
            }
        }

        /// <summary>
        /// 对当前所有已打开窗口应用标题栏主题
        /// </summary>
        public static void ApplyToAllWindows()
        {
            foreach (Window w in Application.Current.Windows)
            {
                ApplyToWindow(w);
            }
        }
    }
}

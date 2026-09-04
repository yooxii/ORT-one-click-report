using Microsoft.Extensions.DependencyInjection;
using ORT一键报告.Main.Views;
using System;
using System.Linq;
using System.Windows;

namespace ORT一键报告.Services
{
    /// <summary>
    /// Toast 提示服务：在“当前聚焦窗口”的指定角落（设置可配，默认右上角）弹出轻量提示，
    /// 数秒后自动淡出。用于替代部分非阻塞式提醒（如下拉修改提示、报告路径未设置提醒）。
    /// </summary>
    public static class ToastService
    {
        private static WindowToast _current;

        /// <summary>
        /// 显示一条 Toast（线程安全，自动切回 UI 线程）
        /// </summary>
        public static void Show(string message, ToastType type = ToastType.Info)
        {
            if (Application.Current == null)
            {
                return;
            }
            if (Application.Current.Dispatcher.CheckAccess())
            {
                ShowCore(message, type);
            }
            else
            {
                Application.Current.Dispatcher.Invoke(() => ShowCore(message, type));
            }
        }

        private static void ShowCore(string message, ToastType type)
        {
            try
            {
                // 替换上一条未消失的 Toast
                if (_current != null)
                {
                    _current.Close();
                    _current = null;
                }

                Window owner = Application.Current.Windows.OfType<Window>()
                    .FirstOrDefault(w => w.IsActive && w is not WindowToast)
                    ?? Application.Current.MainWindow;

                var toast = new WindowToast(message, type);
                toast.PositionNear(owner, GetPosition());
                toast.BeginDisplay();
                _current = toast;
            }
            catch (Exception ex)
            {
                NLog.LogManager.GetCurrentClassLogger().Warn(ex, "显示 Toast 失败");
            }
        }

        /// <summary>
        /// 报告路径（设置→路径）为空时弹出提醒 Toast，供与报告相关的菜单项调用
        /// </summary>
        public static void WarnIfReportPathEmpty()
        {
            try
            {
                var settings = App.ServiceProvider?.GetService<AppSettingsService>();
                if (settings?.Settings?.Paths != null
                    && string.IsNullOrWhiteSpace(settings.Settings.Paths.ReportPath))
                {
                    Show(LanguageService.Get("Toast_ReportPathEmpty"), ToastType.Warning);
                }
            }
            catch
            {
                // 忽略：提醒失败不应阻断主流程
            }
        }

        /// <summary>
        /// 读取设置中的 Toast 位置（默认右上角）
        /// </summary>
        private static string GetPosition()
        {
            try
            {
                return App.ServiceProvider?.GetService<AppSettingsService>()?.Settings?.UI?.ToastPosition
                    ?? "TopRight";
            }
            catch
            {
                return "TopRight";
            }
        }
    }
}

using NLog;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ORT一键报告.Utils
{
    /// <summary>
    /// 内存清理：主窗口最小化到后台（托盘）时调用，把不再需要的托管对象回收掉，
    /// 并把工作集还给系统——程序挂后台时不会一直占着前台运行时用到的内存。
    /// 只回收内存，不断开数据库、不清缓存，恢复窗口后功能照旧（首次操作可能稍慢）。
    /// </summary>
    public static class MemoryTrimmer
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        [DllImport("kernel32.dll")]
        private static extern bool SetProcessWorkingSetSize(IntPtr process, IntPtr minimumWorkingSetSize, IntPtr maximumWorkingSetSize);

        /// <summary>
        /// 最后一次清理后的工作集（KB）；没清理过为 0。用于日志与设置界面确认效果。
        /// </summary>
        public static long LastWorkingSetKb { get; private set; }

        /// <summary>
        /// 执行一次内存清理：整代回收 + 把工作集还给系统。任何一步失败都只记日志，不抛异常。
        /// </summary>
        public static void Trim(string reason = null)
        {
            long before = WorkingSetKb();
            try
            {
                GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, blocking: true, compacting: true);
                GC.WaitForPendingFinalizers();
                // 终结器可能又释放出一批对象，再收一次
                GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, blocking: true, compacting: true);
                SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle, new IntPtr(-1), new IntPtr(-1));
            }
            catch (Exception ex)
            {
                _logger.Warn($"内存清理失败（{reason ?? "未说明"}）: {ex.Message}");
                return;
            }
            LastWorkingSetKb = WorkingSetKb();
            _logger.Info($"内存清理完成（{reason ?? "未说明"}）：工作集 {before} KB → {LastWorkingSetKb} KB");
        }

        /// <summary>
        /// 当前进程工作集（KB）；取不到返回 0
        /// </summary>
        public static long WorkingSetKb()
        {
            try
            {
                Process process = Process.GetCurrentProcess();
                process.Refresh();
                return process.WorkingSet64 / 1024;
            }
            catch (Exception ex)
            {
                _logger.Warn($"读取工作集失败: {ex.Message}");
                return 0;
            }
        }
    }
}

using NLog;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Threading;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 闲置自动执行报告扫描：设置里开启后，每 30 秒检查一次，
    /// 满足「开关开启 + 电脑空闲达到设定时长 + 当前无高耗时任务在跑」时在后台触发一次扫描。
    /// 与 PlanIndexScheduler 同构；扫描本身走 HighCostTaskCoordinator，与计划索引/一键报告互斥。
    /// </summary>
    public class ReportScanScheduler : IDisposable
    {
        /// <summary>闲置自动扫描开关（app_settings 键）</summary>
        public const string SettingAutoKey = "report.scan.auto";

        /// <summary>闲置多少秒后执行（app_settings 键）</summary>
        public const string SettingIdleSecondsKey = "report.scan.idleSeconds";

        /// <summary>默认闲置秒数</summary>
        public const int DefaultIdleSeconds = 300;

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DispatcherTimer _timer = new() { Interval = TimeSpan.FromSeconds(30) };
        private readonly ReportScanService _scan;
        private readonly HighCostTaskCoordinator _coordinator;
        private readonly AppSettingsService _settings;
        private bool _started;

        public ReportScanScheduler(ReportScanService scan, HighCostTaskCoordinator coordinator, AppSettingsService settings)
        {
            _scan = scan;
            _coordinator = coordinator;
            _settings = settings;
        }

        /// <summary>启动定时检查（首轮延迟一点，避免和程序启动/其它后台任务抢资源）</summary>
        public void Start()
        {
            if (_started)
            {
                return;
            }
            _started = true;
            _timer.Tick += async (s, e) => await TickAsync().ConfigureAwait(true);
            _timer.Start();
            _logger.Info($"闲置报告扫描已启用：开关键 {SettingAutoKey}，闲置秒键 {SettingIdleSecondsKey}");
        }

        private async System.Threading.Tasks.Task TickAsync()
        {
            try
            {
                if (!_settings.GetBool(SettingAutoKey, true))
                {
                    return;
                }
                if (_coordinator.IsBusy || _scan.IsRunning)
                {
                    return;
                }
                int idleSeconds = _settings.GetInt(SettingIdleSecondsKey, DefaultIdleSeconds);
                if (idleSeconds < 10)
                {
                    idleSeconds = 10; // 下限保护，避免设置成 0 后一直触发
                }
                if (GetIdleTime() < TimeSpan.FromSeconds(idleSeconds))
                {
                    return;
                }
                _logger.Info($"闲置 {idleSeconds} 秒达到，自动触发报告扫描");
                await _coordinator.RequestStartAsync("报告扫描",
                    cts => _scan.RunAsync(cts.Token)).ConfigureAwait(true);
            }
            catch (Exception ex)
            {
                _logger.Warn($"闲置自动报告扫描失败: {ex.Message}");
            }
        }

        /// <summary>系统空闲时长（无人操作键盘鼠标）</summary>
        public static TimeSpan GetIdleTime()
        {
            try
            {
                LASTINPUTINFO info = new() { cbSize = (uint)Marshal.SizeOf(typeof(LASTINPUTINFO)) };
                if (!GetLastInputInfo(ref info))
                {
                    return TimeSpan.Zero;
                }
                uint now = unchecked((uint)Environment.TickCount);
                uint elapsed = now - info.dwTime;
                return TimeSpan.FromMilliseconds(elapsed);
            }
            catch
            {
                return TimeSpan.Zero;
            }
        }

        public void Dispose()
        {
            _timer.Stop();
            _started = false;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct LASTINPUTINFO
        {
            public uint cbSize;
            public uint dwTime;
        }

        [DllImport("user32.dll")]
        private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);
    }
}

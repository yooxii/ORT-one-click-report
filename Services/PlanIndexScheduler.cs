using NLog;
using ORT一键报告.Models;
using System;
using System.Runtime.InteropServices;
using System.Windows.Threading;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 空闲自动执行计划索引：设置里开启后，每分钟检查一次，
    /// 满足「已登录管理员 + 电脑空闲达到设定时长 + 当前无人执行任务」时在后台接着建立计划索引。
    /// 因为任务与明细都落库，所以关掉程序/换台电脑都会从上次的断点继续。
    /// </summary>
    public class PlanIndexScheduler : IDisposable
    {
        /// <summary>空闲自动索引开关（app_settings 键）</summary>
        public const string SettingAutoKey = "plan.index.auto";

        /// <summary>空闲多久后执行（分钟，app_settings 键）</summary>
        public const string SettingIdleMinutesKey = "plan.index.idleMinutes";

        /// <summary>默认空闲分钟数</summary>
        public const int DefaultIdleMinutes = 10;

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DispatcherTimer _timer = new() { Interval = TimeSpan.FromMinutes(1) };
        private readonly PlanIndexService _index;
        private readonly AppSettingsService _settings;
        private readonly IPermissionService _permission;
        private bool _started;

        public PlanIndexScheduler(PlanIndexService index, AppSettingsService settings, IPermissionService permission)
        {
            _index = index;
            _settings = settings;
            _permission = permission;
        }

        /// <summary>自动执行完成后的通知（供界面提示用户）</summary>
        public event Action<string> Finished;

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
            _logger.Info($"空闲计划索引已启用：开关键 {SettingAutoKey}，空闲分钟键 {SettingIdleMinutesKey}");
        }

        /// <summary>一次检查：条件满足就在后台执行（互斥由 PlanIndexService 的认领机制保证）</summary>
        private async System.Threading.Tasks.Task TickAsync()
        {
            try
            {
                if (_index.IsRunning || !_settings.GetBool(SettingAutoKey, false))
                {
                    return;
                }
                // 自动执行只面向管理员：避免普通用户登录时后台消耗数据库与磁盘
                if (!_permission.Can("admin.manage"))
                {
                    return;
                }
                int idleMinutes = Math.Max(1, _settings.GetInt(SettingIdleMinutesKey, DefaultIdleMinutes));
                if (GetIdleTime() < TimeSpan.FromMinutes(idleMinutes))
                {
                    return;
                }
                string root = _settings.ReportDir;
                if (string.IsNullOrWhiteSpace(root))
                {
                    return;
                }
                PlanIndexRunResult result = await _index.RunAsync(root, _permission.CurrentUser).ConfigureAwait(true);
                if (result.Started)
                {
                    _logger.Info($"空闲自动建立计划索引：{result.Message}");
                    Finished?.Invoke(result.Message);
                }
                else if (!string.IsNullOrEmpty(result.Message))
                {
                    _logger.Info($"空闲自动索引跳过：{result.Message}");
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"空闲自动建立计划索引失败: {ex.Message}");
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

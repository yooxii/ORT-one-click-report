using NLog;
using System;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 定时备份调度：启动后先查一次，之后每 30 分钟查一次，按策略执行——
    /// 距上次全量满 7 天（或还没有全量）做一次全量备份，否则当天还没做增量就做一次增量备份。
    /// 备份在后台线程执行（网络共享上的库可能较大，不能卡界面）；
    /// 多台电脑同时到点时由备份目录里的锁文件互斥，只有一台真正执行。
    /// </summary>
    public class DatabaseBackupScheduler : IDisposable
    {
        /// <summary>启动后第一次检查的延迟（先让程序启动、登录、加载数据跑完）</summary>
        private static readonly TimeSpan FirstCheckDelay = TimeSpan.FromSeconds(90);

        /// <summary>之后的检查间隔</summary>
        private static readonly TimeSpan CheckInterval = TimeSpan.FromMinutes(30);

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DispatcherTimer _timer = new();
        private readonly DatabaseBackupService _backup;
        private bool _started;
        private bool _running;

        public DatabaseBackupScheduler(DatabaseBackupService backup)
        {
            _backup = backup;
        }

        /// <summary>后台生成了备份时通知界面（消息可直接显示）</summary>
        public event Action<string> Finished;

        /// <summary>启动定时检查</summary>
        public void Start()
        {
            if (_started)
            {
                return;
            }
            _started = true;
            _timer.Interval = FirstCheckDelay;
            _timer.Tick += OnTick;
            _timer.Start();
            _logger.Info($"定时备份已启用：间隔 {CheckInterval.TotalMinutes:0} 分钟（首次 {FirstCheckDelay.TotalSeconds:0} 秒后检查），策略＝每周一次全量 + 每天一次增量");
        }

        private async void OnTick(object sender, EventArgs e)
        {
            // 第一次之后按正常间隔检查
            _timer.Interval = CheckInterval;
            if (_running)
            {
                return;
            }
            _running = true;
            try
            {
                DatabaseBackupResult result = await Task.Run(() => _backup.RunScheduled()).ConfigureAwait(true);
                if (result.Created)
                {
                    _logger.Info($"定时备份：{result.Message}");
                    Finished?.Invoke(result.Message);
                }
                else if (!result.Success)
                {
                    _logger.Warn($"定时备份未完成：{result.Message}");
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"定时备份失败: {ex.Message}");
            }
            finally
            {
                _running = false;
            }
        }

        public void Dispose()
        {
            _timer.Stop();
            _timer.Tick -= OnTick;
            _started = false;
        }
    }
}

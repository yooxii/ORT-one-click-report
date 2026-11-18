using NLog;
using ORT一键报告.Models;
using System;
using System.Globalization;
using System.Windows.Threading;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 预领的更新提醒：
    /// 1) 月排程：每周固定时间提醒导入最新排程（排程会随时调整，机种可能挪到别的日期）；
    /// 2) ORT看板：每天固定时间提醒导入最新看板（看板天天刷新，跨月后整份都会变）。
    /// 只在本机弹提示（Toast），提醒时间与「今天是否已提醒」都记在本机设置里；
    /// 程序启动时若已过提醒时间且今天还没提醒过，会立刻补一次。
    /// </summary>
    public class PreIssueReminderScheduler : IDisposable
    {
        /// <summary>提醒种类：月排程（每周）</summary>
        public const string KindSchedule = "schedule";

        /// <summary>提醒种类：ORT看板（每天）</summary>
        public const string KindBoard = "board";

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DispatcherTimer _timer = new() { Interval = TimeSpan.FromMinutes(1) };
        private bool _started;

        /// <summary>当前时间来源（测试可替换）</summary>
        internal Func<DateTime> Now { get; set; } = () => DateTime.Now;

        /// <summary>到点提醒时触发（参数：提醒种类、提醒文本）</summary>
        public event Action<string, string> Prompted;

        /// <summary>启动定时检查</summary>
        public void Start()
        {
            if (_started)
            {
                return;
            }
            _started = true;
            _timer.Tick += (s, e) => Tick();
            _timer.Start();
            // 启动后先检查一次（错过的时间点在这里补上）
            Tick();
            _logger.Info("预领导入提醒已启用（每周排程 / 每天看板）");
        }

        /// <summary>一次检查：到了提醒时间且今天还没提醒过，就提醒一次</summary>
        internal void Tick()
        {
            try
            {
                LocalSettings settings = LocalSettingsStore.Read();
                PreIssueReminderSettings reminder = settings.PreIssueReminders ?? new PreIssueReminderSettings();
                DateTime now = Now();
                string today = now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                bool weeklyDue = IsWeeklyDue(reminder, now, out DateTime weeklyAt);
                bool dailyDue = IsDailyDue(reminder, now, out DateTime dailyAt);
                if (!weeklyDue && !dailyDue)
                {
                    return;
                }
                LocalSettingsStore.Update(s =>
                {
                    s.PreIssueReminders ??= new PreIssueReminderSettings();
                    if (weeklyDue)
                    {
                        s.PreIssueReminders.LastWeeklyPrompt = today;
                    }
                    if (dailyDue)
                    {
                        s.PreIssueReminders.LastDailyPrompt = today;
                    }
                });
                if (weeklyDue)
                {
                    _logger.Info($"月排程更新提醒（计划时间 {weeklyAt:yyyy-MM-dd HH:mm}）");
                    Prompted?.Invoke(KindSchedule, LanguageService.Get("PreIssue_Reminder_Weekly"));
                }
                if (dailyDue)
                {
                    _logger.Info($"ORT看板更新提醒（计划时间 {dailyAt:yyyy-MM-dd HH:mm}）");
                    Prompted?.Invoke(KindBoard, LanguageService.Get("PreIssue_Reminder_Daily"));
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"预领更新提醒检查失败: {ex.Message}");
            }
        }

        /// <summary>每周排程提醒是否到点（今天该提醒、已过时间、今天还没提醒过）</summary>
        internal static bool IsWeeklyDue(PreIssueReminderSettings settings, DateTime now, out DateTime scheduledAt)
        {
            scheduledAt = now;
            if (settings == null || !settings.WeeklyEnabled)
            {
                return false;
            }
            DayOfWeek day = ParseDayOfWeek(settings.WeeklyDay, DayOfWeek.Monday);
            if (now.DayOfWeek != day)
            {
                return false;
            }
            TimeSpan time = ParseTime(settings.WeeklyTime, new TimeSpan(8, 30, 0));
            scheduledAt = now.Date + time;
            if (now < scheduledAt)
            {
                return false;
            }
            return !string.Equals(settings.LastWeeklyPrompt, now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), StringComparison.Ordinal);
        }

        /// <summary>每天看板提醒是否到点</summary>
        internal static bool IsDailyDue(PreIssueReminderSettings settings, DateTime now, out DateTime scheduledAt)
        {
            scheduledAt = now;
            if (settings == null || !settings.DailyEnabled)
            {
                return false;
            }
            TimeSpan time = ParseTime(settings.DailyTime, new TimeSpan(8, 30, 0));
            scheduledAt = now.Date + time;
            if (now < scheduledAt)
            {
                return false;
            }
            return !string.Equals(settings.LastDailyPrompt, now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), StringComparison.Ordinal);
        }

        /// <summary>解析 HH:mm（认不出用默认值）</summary>
        public static TimeSpan ParseTime(string text, TimeSpan fallback)
        {
            if (!string.IsNullOrWhiteSpace(text)
                && TimeSpan.TryParseExact(text.Trim(), @"h\:mm", CultureInfo.InvariantCulture, out TimeSpan time)
                && time >= TimeSpan.Zero && time < TimeSpan.FromDays(1))
            {
                return time;
            }
            return fallback;
        }

        /// <summary>解析星期名（Monday..Sunday，认不出用默认值）</summary>
        public static DayOfWeek ParseDayOfWeek(string text, DayOfWeek fallback)
        {
            if (!string.IsNullOrWhiteSpace(text)
                && Enum.TryParse(text.Trim(), true, out DayOfWeek day)
                && Enum.IsDefined(typeof(DayOfWeek), day))
            {
                return day;
            }
            return fallback;
        }

        public void Dispose()
        {
            _timer.Stop();
            _started = false;
        }
    }
}

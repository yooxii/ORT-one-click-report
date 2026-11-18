using ORT一键报告.Models;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// 预领的更新提醒设置（本机保存）：
    /// 月排程每周固定时间提醒导入，ORT 看板每天固定时间提醒导入；同一天只提醒一次。
    /// </summary>
    public partial class WindowPreIssueReminder : Window
    {
        /// <summary>星期下拉的候选（值 ↔ 显示名）</summary>
        private static readonly (DayOfWeek Day, string Key)[] Days =
        [
            (DayOfWeek.Monday, "PreIssue_Day_Monday"),
            (DayOfWeek.Tuesday, "PreIssue_Day_Tuesday"),
            (DayOfWeek.Wednesday, "PreIssue_Day_Wednesday"),
            (DayOfWeek.Thursday, "PreIssue_Day_Thursday"),
            (DayOfWeek.Friday, "PreIssue_Day_Friday"),
            (DayOfWeek.Saturday, "PreIssue_Day_Saturday"),
            (DayOfWeek.Sunday, "PreIssue_Day_Sunday")
        ];

        public WindowPreIssueReminder()
        {
            InitializeComponent();
            PreIssueReminderSettings settings = LocalSettingsStore.Read().PreIssueReminders ?? new PreIssueReminderSettings();

            List<DayOption> options = [];
            foreach ((DayOfWeek day, string key) in Days)
            {
                options.Add(new DayOption { Day = day, Text = LanguageService.Get(key) });
            }
            cb_day.ItemsSource = options;
            cb_day.SelectedItem = options.Find(o => o.Day == PreIssueReminderScheduler.ParseDayOfWeek(settings.WeeklyDay, DayOfWeek.Monday));

            chk_weekly.IsChecked = settings.WeeklyEnabled;
            chk_daily.IsChecked = settings.DailyEnabled;
            txt_weekly_time.Text = settings.WeeklyTime ?? "08:30";
            txt_daily_time.Text = settings.DailyTime ?? "08:30";
            txt_tip.Text = string.Format(LanguageService.Get("PreIssue_ReminderLast"),
                string.IsNullOrWhiteSpace(settings.LastWeeklyPrompt) ? "-" : settings.LastWeeklyPrompt,
                string.IsNullOrWhiteSpace(settings.LastDailyPrompt) ? "-" : settings.LastDailyPrompt);
        }

        private void Btn_Save_Click(object sender, RoutedEventArgs e)
        {
            string weeklyTime = (txt_weekly_time.Text ?? "").Trim();
            string dailyTime = (txt_daily_time.Text ?? "").Trim();
            if (!IsValidTime(weeklyTime) || !IsValidTime(dailyTime))
            {
                _ = MessageBox.Show(LanguageService.Get("PreIssue_ReminderTimeInvalid"), Title,
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            DayOfWeek day = cb_day.SelectedItem is DayOption option ? option.Day : DayOfWeek.Monday;
            LocalSettingsStore.Update(s =>
            {
                s.PreIssueReminders ??= new PreIssueReminderSettings();
                s.PreIssueReminders.WeeklyEnabled = chk_weekly.IsChecked == true;
                s.PreIssueReminders.DailyEnabled = chk_daily.IsChecked == true;
                s.PreIssueReminders.WeeklyDay = day.ToString();
                s.PreIssueReminders.WeeklyTime = weeklyTime;
                s.PreIssueReminders.DailyTime = dailyTime;
            });
            DialogResult = true;
        }

        private static bool IsValidTime(string text) =>
            TimeSpan.TryParseExact(text, @"h\:mm", CultureInfo.InvariantCulture, out TimeSpan time)
            && time >= TimeSpan.Zero && time < TimeSpan.FromDays(1);

        /// <summary>星期下拉项</summary>
        private sealed class DayOption
        {
            public DayOfWeek Day { get; set; }
            public string Text { get; set; }
            public override string ToString() => Text;
        }
    }
}

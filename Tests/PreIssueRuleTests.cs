using NUnit.Framework;
using ORT一键报告.Models;
using ORT一键报告.Services;
using System;

namespace ORT一键报告.Tests
{
    /// <summary>
    /// 预领（月排程 / ORT看板 / 准备领用）的解析与提醒规则测试：
    /// 「機種/數量」拆分、月份识别、单号里的年月、SIM87 的 Y 判定、每周/每天提醒是否到点。
    /// </summary>
    [TestFixture]
    public class PreIssueRuleTests
    {
        /* ------------------ 「機種/數量」拆分 ------------------ */

        [TestCase("OTD026-00BG/708", "OTD026-00BG", 708)]
        [TestCase("FSG032-A2AG/0", "FSG032-A2AG", 0)]
        [TestCase("FSE052-GAAG/1792", "FSE052-GAAG", 1792)]
        [TestCase("WAS003-F90GB/56", "WAS003-F90GB", 56)]
        [TestCase("  FSL020-00AG / 1500  ", "FSL020-00AG", 1500)]
        [TestCase("ABC123", "ABC123", 0)]
        [TestCase("", "", 0)]
        public void ParseModelQty_SplitsModelAndQuantity(string text, string model, int qty)
        {
            (string actualModel, int actualQty) = PreIssueService.ParseModelQty(text);
            Assert.That(actualModel, Is.EqualTo(model));
            Assert.That(actualQty, Is.EqualTo(qty));
        }

        /* ------------------ 月份识别 ------------------ */

        [TestCase("202611", "202611")]
        [TestCase("2026-11", "202611")]
        [TestCase("2026/11", "202611")]
        [TestCase("2026年11月", "202611")]
        [TestCase("  202612  ", "202612")]
        [TestCase("202613", null)]
        [TestCase("", null)]
        [TestCase("排程", null)]
        public void NormalizeMonth_AcceptsCommonFormats(string text, string expected)
        {
            Assert.That(PreIssueService.NormalizeMonth(text), Is.EqualTo(expected));
        }

        [TestCase("20261117", "202611")]
        [TestCase("2026/11/17", "202611")]
        [TestCase("2026-12-01", "202612")]
        [TestCase("", null)]
        [TestCase("abc", null)]
        public void MonthFromYmd_ReadsBoardDate(string text, string expected)
        {
            Assert.That(PreIssueService.MonthFromYmd(text), Is.EqualTo(expected));
        }

        [TestCase("T01261100166", "202611")]
        [TestCase("T01261001009", "202610")]
        [TestCase("T01250100001", "202501")]
        [TestCase("", null)]
        [TestCase("XXXX", null)]
        public void MonthFromVchno_ReadsVoucherNumber(string text, string expected)
        {
            Assert.That(PreIssueService.MonthFromVchno(text), Is.EqualTo(expected));
        }

        /* ------------------ SIM87 / 状态判定 ------------------ */

        [TestCase("Y", true)]
        [TestCase("y", true)]
        [TestCase("是", true)]
        [TestCase("N", false)]
        [TestCase("n", false)]
        [TestCase("", false)]
        [TestCase(null, false)]
        public void IsYes_OnlyAcceptsYesMarks(string text, bool expected)
        {
            Assert.That(PreIssueCodes.IsYes(text), Is.EqualTo(expected));
        }

        [Test]
        public void IsPendingState_TreatsEmptyAsPending()
        {
            Assert.That(PreIssueCodes.IsPendingState(null), Is.True);
            Assert.That(PreIssueCodes.IsPendingState(""), Is.True);
            Assert.That(PreIssueCodes.IsPendingState(PreIssueCodes.PendingState), Is.True);
            Assert.That(PreIssueCodes.IsPendingState(PreIssueCodes.OpenedState), Is.False);
        }

        /* ------------------ 每周 / 每天提醒是否到点 ------------------ */

        [Test]
        public void WeeklyDue_FiresOnConfiguredDayAfterTime()
        {
            PreIssueReminderSettings settings = new()
            {
                WeeklyEnabled = true,
                WeeklyDay = "Monday",
                WeeklyTime = "08:30"
            };
            // 2026-11-16 是周一
            DateTime monday = new(2026, 11, 16, 9, 0, 0);
            Assert.That(PreIssueReminderScheduler.IsWeeklyDue(settings, monday, out DateTime at), Is.True);
            Assert.That(at, Is.EqualTo(new DateTime(2026, 11, 16, 8, 30, 0)));
        }

        [Test]
        public void WeeklyDue_NotBeforeTimeOrOtherDay()
        {
            PreIssueReminderSettings settings = new() { WeeklyDay = "Monday", WeeklyTime = "08:30" };
            DateTime early = new(2026, 11, 16, 8, 0, 0);
            Assert.That(PreIssueReminderScheduler.IsWeeklyDue(settings, early, out _), Is.False);

            DateTime tuesday = new(2026, 11, 17, 9, 0, 0);
            Assert.That(PreIssueReminderScheduler.IsWeeklyDue(settings, tuesday, out _), Is.False);
        }

        [Test]
        public void WeeklyDue_OnlyOncePerDay()
        {
            PreIssueReminderSettings settings = new() { WeeklyDay = "Monday", WeeklyTime = "08:30" };
            DateTime monday = new(2026, 11, 16, 9, 0, 0);
            settings.LastWeeklyPrompt = "2026-11-16";
            Assert.That(PreIssueReminderScheduler.IsWeeklyDue(settings, monday, out _), Is.False);
        }

        [Test]
        public void WeeklyDue_DisabledNeverFires()
        {
            PreIssueReminderSettings settings = new() { WeeklyEnabled = false, WeeklyDay = "Monday", WeeklyTime = "08:30" };
            Assert.That(PreIssueReminderScheduler.IsWeeklyDue(settings, new DateTime(2026, 11, 16, 9, 0, 0), out _), Is.False);
        }

        [Test]
        public void DailyDue_FiresAfterTimeOncePerDay()
        {
            PreIssueReminderSettings settings = new() { DailyEnabled = true, DailyTime = "08:00" };
            DateTime ten = new(2026, 11, 18, 10, 0, 0);
            Assert.That(PreIssueReminderScheduler.IsDailyDue(settings, ten, out DateTime at), Is.True);
            Assert.That(at, Is.EqualTo(new DateTime(2026, 11, 18, 8, 0, 0)));

            settings.LastDailyPrompt = "2026-11-18";
            Assert.That(PreIssueReminderScheduler.IsDailyDue(settings, ten, out _), Is.False);

            DateTime seven = new(2026, 11, 19, 7, 0, 0);
            settings.LastDailyPrompt = null;
            Assert.That(PreIssueReminderScheduler.IsDailyDue(settings, seven, out _), Is.False);
        }

        [TestCase("08:30", 8, 30)]
        [TestCase("23:59", 23, 59)]
        [TestCase("8:30", 8, 30)]
        public void ParseTime_ReadsHourMinute(string text, int hour, int minute)
        {
            Assert.That(PreIssueReminderScheduler.ParseTime(text, TimeSpan.FromHours(1)),
                Is.EqualTo(new TimeSpan(hour, minute, 0)));
        }

        [TestCase("乱写")]
        [TestCase("25:00")]
        [TestCase("")]
        public void ParseTime_FallsBackOnInvalidText(string text)
        {
            TimeSpan fallback = new(8, 30, 0);
            Assert.That(PreIssueReminderScheduler.ParseTime(text, fallback), Is.EqualTo(fallback));
        }

        [Test]
        public void ParseDayOfWeek_ReadsEnglishName()
        {
            Assert.That(PreIssueReminderScheduler.ParseDayOfWeek("Friday", DayOfWeek.Sunday), Is.EqualTo(DayOfWeek.Friday));
            Assert.That(PreIssueReminderScheduler.ParseDayOfWeek("friday", DayOfWeek.Sunday), Is.EqualTo(DayOfWeek.Friday));
            Assert.That(PreIssueReminderScheduler.ParseDayOfWeek("星期五", DayOfWeek.Sunday), Is.EqualTo(DayOfWeek.Sunday));
        }

        /* ------------------ 看板行比对键 ------------------ */

        [Test]
        public void BoardMatchKey_PrefersJobNo()
        {
            OrtBoardRow withJob = new() { JobNo = " AH0027EA0024 ", Model = "ADC027-YEAG", ProduceDate = "20261205" };
            Assert.That(withJob.MatchKey, Is.EqualTo("AH0027EA0024"));

            OrtBoardRow withoutJob = new() { JobNo = null, Model = "ADC027-YEAG", ProduceDate = "20261205" };
            Assert.That(withoutJob.MatchKey, Is.EqualTo("ADC027-YEAG|20261205"));
        }

        [Test]
        public void Stocked_ReactsToStocked87()
        {
            OrtBoardRow row = new() { Stocked87 = "N" };
            Assert.That(row.Stocked, Is.False);
            row.Stocked87 = "Y";
            Assert.That(row.Stocked, Is.True);
        }
    }
}

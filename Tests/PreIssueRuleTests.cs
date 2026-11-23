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

        /* ------------------ 导入前的文件类型校验（导错文件的拦截） ------------------ */

        [Test]
        public void LooksLikeVoucherExport_RecognizesVoucherHeaders()
        {
            string[] voucher = ["VCHNO_CODE", "VCHNO", "VCH_SEQ", "DIVISION_CODE", "WARE_CODE", "PART_NO", "REQ_QTY"];
            Assert.That(PreIssueValidation.LooksLikeVoucherExport(voucher), Is.True);

            string[] schedule = ["年月", "線別", "1日", "2日", "3日"];
            Assert.That(PreIssueValidation.LooksLikeVoucherExport(schedule), Is.False);
        }

        [Test]
        public void LooksLikeSchedule_RequiresMonthLineAndDayColumns()
        {
            Assert.That(PreIssueValidation.LooksLikeSchedule(["年月", "線別", "1日", "2日", "3日"], out string why), Is.True);
            Assert.That(why, Is.Null);

            // 没有日列
            Assert.That(PreIssueValidation.LooksLikeSchedule(["年月", "線別", "A", "B", "C"], out _), Is.False);
            // 第 1 列不是年月（开单数据的表头）
            Assert.That(PreIssueValidation.LooksLikeSchedule(["VCHNO", "PART_NO", "REQ_QTY", "1日", "2日", "3日"], out _), Is.False);
            // 第 2 列不是線別
            Assert.That(PreIssueValidation.LooksLikeSchedule(["年月", "工令", "1日", "2日", "3日"], out _), Is.False);
        }

        [Test]
        public void LooksLikeOrtBoard_RecognizesBoardHeaders()
        {
            string[] board = ["日期", "類別", "機種", "生產線別", "生產日期", "工令", "SIM87入庫否"];
            Assert.That(PreIssueValidation.LooksLikeOrtBoard(board), Is.True);
            Assert.That(PreIssueValidation.LooksLikeOrtBoard(["年月", "線別", "1日"]), Is.False);
        }

        [TestCase("1日", 1)]
        [TestCase("15日", 15)]
        [TestCase("31日", 31)]
        [TestCase("01日", 1)]
        public void IsDayHeader_ReadsDay(string text, int expected)
        {
            Assert.That(PreIssueValidation.IsDayHeader(text, out int day), Is.True);
            Assert.That(day, Is.EqualTo(expected));
        }

        [TestCase("16")]
        [TestCase("32日")]
        [TestCase("0日")]
        [TestCase("日")]
        [TestCase("")]
        [TestCase("VCHNO")]
        public void IsDayHeader_RejectsOthers(string text)
        {
            Assert.That(PreIssueValidation.IsDayHeader(text, out _), Is.False);
        }

        [TestCase("OTD026-00BG/708", true)]
        [TestCase("FSG032-A2AG/0", true)]
        [TestCase("  FSL020-00AG / 1500  ", true)]
        [TestCase("T", false)]
        [TestCase("8CT1", false)]
        [TestCase("H", false)]
        [TestCase("ADL030-YSAG", false)]
        [TestCase("ABC/abc", false)]
        [TestCase("ABC/", false)]
        [TestCase("", false)]
        public void LooksLikeModelQty_ChecksCellShape(string text, bool expected)
        {
            Assert.That(PreIssueValidation.LooksLikeModelQty(text), Is.EqualTo(expected));
        }

        /// <summary>
        /// 回归：把开单数据（VCHNO … PART_NO）误当月排程导入时，
        /// 表头校验必须直接判否——这就是「提前拦截」的核心判定。
        /// </summary>
        [Test]
        public void VoucherFile_CannotPassAsSchedule()
        {
            string[] voucherHeaders =
            [
                "VCHNO_CODE", "VCHNO", "VCH_SEQ", "DIVISION_CODE", "WARE_CODE", "TR_CODE",
                "PART_NO", "REQ_QTY", "IVER_ID", "VCH_STATUS", "REQ_DEPT"
            ];
            Assert.That(PreIssueValidation.LooksLikeSchedule(voucherHeaders, out string why), Is.False);
            Assert.That(why, Is.Not.Null.And.Not.Empty);
            Assert.That(PreIssueValidation.LooksLikeVoucherExport(voucherHeaders), Is.True);
        }

        /* ------------------ 黑名单：精确机种 + 正则模式 ------------------ */

        [Test]
        public void BlacklistMatcher_ExactMatchIgnoresCase()
        {
            OrtBlacklistMatcher matcher = OrtBlacklistMatcher.Create(
            [
                new OrtBlacklistItem { Model = "ADC027-YEAG", IsPattern = false }
            ]);
            Assert.That(matcher.Matches("ADC027-YEAG"), Is.True);
            Assert.That(matcher.Matches("adc027-yeag"), Is.True);
            Assert.That(matcher.Matches("ADC027-YEAG "), Is.True);
            Assert.That(matcher.Matches("ADC027-YEAG2"), Is.False);
            Assert.That(matcher.Count, Is.EqualTo(1));
            Assert.That(matcher.ExactCount, Is.EqualTo(1));
        }

        [Test]
        public void BlacklistMatcher_PatternMatchesSubstring()
        {
            OrtBlacklistMatcher matcher = OrtBlacklistMatcher.Create(
            [
                new OrtBlacklistItem { Model = "FSL0.*", IsPattern = true }
            ]);
            Assert.That(matcher.Matches("FSL020-00AG"), Is.True);
            Assert.That(matcher.Matches("fsl031-243g"), Is.True);
            Assert.That(matcher.Matches("ADL030-YSAG"), Is.False);
            Assert.That(matcher.PatternCount, Is.EqualTo(1));
        }

        [Test]
        public void BlacklistMatcher_PatternCanAnchorWholeString()
        {
            OrtBlacklistMatcher matcher = OrtBlacklistMatcher.Create(
            [
                new OrtBlacklistItem { Model = "^WAS003-F9\\dGB$", IsPattern = true }
            ]);
            Assert.That(matcher.Matches("WAS003-F90GB"), Is.True);
            Assert.That(matcher.Matches("WAS003-F90GB-1"), Is.False);
        }

        [Test]
        public void BlacklistMatcher_ReturnsMatchedRuleForHint()
        {
            OrtBlacklistMatcher matcher = OrtBlacklistMatcher.Create(
            [
                new OrtBlacklistItem { Model = "FSL0.*", IsPattern = true }
            ]);
            OrtBlacklistRule rule = matcher.Match("FSL020-00AG");
            Assert.That(rule, Is.Not.Null);
            Assert.That(rule.IsPattern, Is.True);
            Assert.That(OrtBlacklistMatcher.Describe(rule), Is.EqualTo("正则 FSL0.*"));
            Assert.That(matcher.Match("ADL030-YSAG"), Is.Null);
        }

        [Test]
        public void BlacklistMatcher_InvalidPatternIsRejectedOrSkipped()
        {
            OrtBlacklistMatcher matcher = new();
            Assert.That(matcher.Add("[", true), Is.False);   // 缺右括号的正则
            Assert.That(matcher.Count, Is.EqualTo(0));

            Assert.That(OrtBlacklistMatcher.TryCompile("[", out _, out string error), Is.False);
            Assert.That(error, Is.Not.Null.And.Not.Empty);

            // 从库里读出来时写错的条目被跳过，不影响其它规则
            OrtBlacklistMatcher mixed = OrtBlacklistMatcher.Create(
            [
                new OrtBlacklistItem { Model = "[", IsPattern = true },
                new OrtBlacklistItem { Model = "OK.*", IsPattern = true }
            ]);
            Assert.That(mixed.PatternCount, Is.EqualTo(1));
            Assert.That(mixed.Matches("OK123"), Is.True);
        }

        [Test]
        public void BlacklistMatcher_EmptyArgumentsAreIgnored()
        {
            OrtBlacklistMatcher matcher = OrtBlacklistMatcher.Create(
            [
                new OrtBlacklistItem { Model = null, IsPattern = false },
                new OrtBlacklistItem { Model = "   ", IsPattern = false }
            ]);
            Assert.That(matcher.Count, Is.EqualTo(0));
            Assert.That(matcher.Matches(null), Is.False);
            Assert.That(matcher.Matches(""), Is.False);
            Assert.That(matcher.Match("ABC"), Is.Null);
        }

        /* ------------------ 机种 → 同系列正则 ------------------ */

        [Test]
        public void SeriesPattern_BuildsFromSixPlusFourModel()
        {
            Assert.That(OrtBlacklistMatcher.TryBuildSeriesPattern("FSL020-00AG", out string pattern,
                out string series, out string customer, out string version), Is.True);
            Assert.That(pattern, Is.EqualTo("^FSL020-[A-Za-z0-9]+$"));
            Assert.That(series, Is.EqualTo("FSL020"));
            Assert.That(customer, Is.EqualTo("00"));
            Assert.That(version, Is.EqualTo("AG"));
        }

        /// <summary>后段 5 位的现场机种（如 WAS003-F90GB）：按系列生成正则，但不拆客户短码/版本</summary>
        [Test]
        public void SeriesPattern_WorksForFiveCharSuffixWithoutCustomerVersion()
        {
            Assert.That(OrtBlacklistMatcher.TryBuildSeriesPattern("WAS003-F90GB", out string pattern,
                out string series, out string customer, out string version), Is.True);
            Assert.That(pattern, Is.EqualTo("^WAS003-[A-Za-z0-9]+$"));
            Assert.That(series, Is.EqualTo("WAS003"));
            Assert.That(customer, Is.Null);
            Assert.That(version, Is.Null);

            OrtBlacklistMatcher matcher = OrtBlacklistMatcher.Create(
            [
                new OrtBlacklistItem { Model = pattern, IsPattern = true }
            ]);
            Assert.That(matcher.Matches("WAS003-F90GK"), Is.True);
            Assert.That(matcher.Matches("WAS003-F91GB"), Is.True);
            Assert.That(matcher.Matches("WAR008-F91GF"), Is.False);
        }

        [Test]
        public void SeriesPattern_GeneratedRuleMatchesWholeSeriesOnly()
        {
            OrtBlacklistMatcher.TryBuildSeriesPattern("FSL020-00AG", out string pattern, out _, out _, out _);
            OrtBlacklistMatcher matcher = OrtBlacklistMatcher.Create(
            [
                new OrtBlacklistItem { Model = pattern, IsPattern = true }
            ]);
            // 同一个系列（前 6 位相同）：客户短码与版本随意
            Assert.That(matcher.Matches("FSL020-00AG"), Is.True);
            Assert.That(matcher.Matches("FSL020-171G"), Is.True);
            Assert.That(matcher.Matches("fsl020-7kag"), Is.True);
            // 别的系列 / 形状不对
            Assert.That(matcher.Matches("FSL021-171G"), Is.False);
            Assert.That(matcher.Matches("FSL0200-171G"), Is.False);
            Assert.That(matcher.Matches("FSL020-171"), Is.True);      // 同系列，后段位数不同也挡
            Assert.That(matcher.Matches("FSL020-171GX"), Is.True);
            Assert.That(matcher.Matches("FSL020"), Is.False);
            Assert.That(matcher.Matches("FSL020_171G"), Is.False);
        }

        [TestCase("FSL020-00AG", "FSL020", "00", "AG")]
        [TestCase("fsl020-7kag", "fsl020", "7k", "ag")]
        public void SeriesParse_HandlesCase(string model, string series, string customer, string version)
        {
            Assert.That(OrtBlacklistMatcher.TryParseSeries(model, out string s, out string c, out string v), Is.True);
            Assert.That(s, Is.EqualTo(series));
            Assert.That(c, Is.EqualTo(customer));
            Assert.That(v, Is.EqualTo(version));
        }

        [TestCase("FSL020")]
        [TestCase("FSL02-00AG")]
        [TestCase("FSL0200-00AG")]
        [TestCase("FSL020-")]
        [TestCase("FSL020_00AG")]
        [TestCase("FSL020-00AG-X")]
        [TestCase("")]
        public void SeriesPattern_RejectsOtherShapes(string model)
        {
            Assert.That(OrtBlacklistMatcher.TryBuildSeriesPattern(model, out _, out _, out _, out _), Is.False);
            Assert.That(OrtBlacklistMatcher.IsSeriesModel(model), Is.False);
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

using NUnit.Framework;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ORT一键报告.Tests
{
    /// <summary>
    /// 报告模板排期规则：
    /// 每项测试占连续的若干天（周期小时数折算，不超过 24 小时算 1 天、超过 24 小时按整天数加一天），
    /// 开始日落在工作日；相邻两个跨天测试连排时，后一项接在前一项的结束日当天（共用边界日），整条排期省一天。
    /// </summary>
    [TestFixture]
    public class ReportTemplateScheduleTests
    {
        /// <summary>排期起点：2026-10-12（周一）</summary>
        private static DateTime Monday
        {
            get { return new DateTime(2026, 10, 12); }
        }

        private static ReportTemplateItem Item(string name, string hours)
        {
            return new ReportTemplateItem { TestItemName = name, PeriodHours = hours };
        }

        /// <summary>按小时数依次排期，返回排好期的测试项</summary>
        private static List<ReportTemplateItem> Schedule(DateTime start, params string[] hours)
        {
            List<ReportTemplateItem> items = hours
                .Select((h, i) => Item("测试项" + (i + 1), h))
                .ToList();
            ReportTemplateService.Schedule(items, start);
            return items;
        }

        private static DateTime StartOf(List<ReportTemplateItem> items, int index)
        {
            Assert.That(items[index].Start, Is.Not.Null, "第 " + (index + 1) + " 项没有排到开始日期");
            return items[index].Start.Value;
        }

        private static DateTime EndOf(List<ReportTemplateItem> items, int index)
        {
            Assert.That(items[index].End, Is.Not.Null, "第 " + (index + 1) + " 项没有排到结束日期");
            return items[index].End.Value;
        }

        [Test]
        public void 周期小时折算天数_跨天测试比整天数多一天()
        {
            Assert.That(Item("空周期", null).PeriodDays, Is.EqualTo(1));
            Assert.That(Item("零周期", "0").PeriodDays, Is.EqualTo(1));
            Assert.That(Item("半天", "12").PeriodDays, Is.EqualTo(1));
            Assert.That(Item("一天", "24").PeriodDays, Is.EqualTo(1));
            Assert.That(Item("两天", "48").PeriodDays, Is.EqualTo(3));
            Assert.That(Item("七天", "168").PeriodDays, Is.EqualTo(8));
            Assert.That(Item("十四天", "336").PeriodDays, Is.EqualTo(15));
        }

        [Test]
        public void 两个跨天测试连排_共用边界日省一天()
        {
            List<ReportTemplateItem> items = Schedule(Monday, "168", "48");

            Assert.That(StartOf(items, 0), Is.EqualTo(Monday));
            Assert.That(EndOf(items, 0), Is.EqualTo(Monday.AddDays(7)));
            // 后一项接在前一项的结束日当天，而不是它的第二天
            Assert.That(StartOf(items, 1), Is.EqualTo(Monday.AddDays(7)));
            Assert.That(EndOf(items, 1), Is.EqualTo(Monday.AddDays(9)));
            // 共用边界日不影响每项自己的天数
            Assert.That(items[0].Days, Is.EqualTo(8));
            Assert.That(items[1].Days, Is.EqualTo(3));
        }

        [Test]
        public void 三个跨天测试连排_每处都省一天()
        {
            List<ReportTemplateItem> items = Schedule(Monday, "168", "48", "72");

            Assert.That(StartOf(items, 0), Is.EqualTo(Monday));
            Assert.That(StartOf(items, 1), Is.EqualTo(Monday.AddDays(7)));
            Assert.That(StartOf(items, 2), Is.EqualTo(Monday.AddDays(9)));
            Assert.That(EndOf(items, 2), Is.EqualTo(Monday.AddDays(12)));
        }

        [Test]
        public void 跨天接当天_不省()
        {
            List<ReportTemplateItem> items = Schedule(Monday, "168", "24");

            Assert.That(EndOf(items, 0), Is.EqualTo(Monday.AddDays(7)));
            Assert.That(StartOf(items, 1), Is.EqualTo(Monday.AddDays(8)));
            Assert.That(EndOf(items, 1), Is.EqualTo(Monday.AddDays(8)));
        }

        [Test]
        public void 当天接跨天_不省()
        {
            List<ReportTemplateItem> items = Schedule(Monday, "24", "48");

            Assert.That(StartOf(items, 0), Is.EqualTo(Monday));
            Assert.That(EndOf(items, 0), Is.EqualTo(Monday));
            Assert.That(StartOf(items, 1), Is.EqualTo(Monday.AddDays(1)));
            Assert.That(EndOf(items, 1), Is.EqualTo(Monday.AddDays(3)));
        }

        [Test]
        public void 两个当天项目_依次各占一天()
        {
            List<ReportTemplateItem> items = Schedule(Monday, "12", "24");

            Assert.That(StartOf(items, 0), Is.EqualTo(Monday));
            Assert.That(StartOf(items, 1), Is.EqualTo(Monday.AddDays(1)));
        }

        [Test]
        public void 前项结束日落在周末_后项顺延到周一()
        {
            // 120 小时 => 6 天，10/12（周一）排到 10/17（周六）
            List<ReportTemplateItem> items = Schedule(Monday, "120", "48");

            Assert.That(EndOf(items, 0), Is.EqualTo(Monday.AddDays(5)));
            Assert.That(StartOf(items, 0).DayOfWeek, Is.EqualTo(DayOfWeek.Monday));
            Assert.That(EndOf(items, 0).DayOfWeek, Is.EqualTo(DayOfWeek.Saturday));
            // 边界日是周六，顺延到下一个周一；此时省不到那一天的便宜，但也不会倒退
            Assert.That(StartOf(items, 1), Is.EqualTo(Monday.AddDays(7)));
        }

        [Test]
        public void 起点落在周末_首项顺延到周一()
        {
            // 2026-10-17 是周六 -> 顺延到 10/19（周一）
            List<ReportTemplateItem> items = Schedule(new DateTime(2026, 10, 17), "24");

            Assert.That(StartOf(items, 0), Is.EqualTo(Monday.AddDays(7)));
            Assert.That(StartOf(items, 0).DayOfWeek, Is.EqualTo(DayOfWeek.Monday));
        }

        [Test]
        public void 连续排期两次_结果一致()
        {
            List<ReportTemplateItem> items = Schedule(Monday, "168", "48", "24");
            DateTime[] firstStart = items.Select(i => i.Start.Value).ToArray();
            DateTime[] firstEnd = items.Select(i => i.End.Value).ToArray();

            ReportTemplateService.Schedule(items, Monday);

            Assert.That(items.Select(i => i.Start.Value).ToArray(), Is.EqualTo(firstStart));
            Assert.That(items.Select(i => i.End.Value).ToArray(), Is.EqualTo(firstEnd));
        }

        [Test]
        public void 没有测试项_不报错()
        {
            ReportTemplateService.Schedule(null, Monday);
            ReportTemplateService.Schedule(new List<ReportTemplateItem>(), Monday);
        }
    }
}

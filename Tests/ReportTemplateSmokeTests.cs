using NUnit.Framework;
using NPOI.SS.UserModel;
using ORT一键报告.Models;
using ORT一键报告.Services;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ORT一键报告.Tests
{
    /// <summary>
    /// 报告模板生成（以已有报告为模板）的端到端冒烟：
    /// 拿一份真实已有报告当输入，在临时目录里跑完整条生成链路，核对
    /// Cover（客户/版本/阶段填写、日期公式保留）、Waterfall（跨天排期、序列号/工令）与
    /// TestStatus（每项测试 total 列置空）。
    ///
    /// 这是一条**只读**的显式用例：不改动任何已有报告，产物全部落在临时目录并在结束时删除；
    /// 需要手工触发（默认不随测试集运行），找不到样本报告时自动跳过。
    /// </summary>
    [TestFixture]
    [Explicit("端到端冒烟，需要真实的已有报告样本，按需手工运行")]
    public class ReportTemplateSmokeTests
    {
        /// <summary>样本报告所在目录（D:\Documents\ORT 下任意一份已有报告）</summary>
        private const string SampleRoot = @"D:\Documents\ORT";
        private const string SampleFolder = "FSJ017-EL2G ORT Test Report (WK2616)_RT260427";
        private const string Model = "FSJ017-EL2G";

        /// <summary>本次排期用的测试项（小时数）：热冲击 24h、EMI 24h、Burn In 168h（跨 8 天）</summary>
        private static readonly (string Name, string Category, string Hours)[] Plan =
        [
            ("Thermal Shock Test", "RELIABILITY TEST", "24"),
            ("Conducted EMI Measurement", "EMC", "24"),
            ("Burn In Test", "RELIABILITY TEST", "168")
        ];

        [Test]
        public void 已有报告作模板_生成结果符合预期()
        {
            string sample = Path.Combine(SampleRoot, SampleFolder);
            if (!Directory.Exists(sample))
            {
                Assert.Ignore($"找不到样本报告：{sample}");
            }

            string work = Path.Combine(Path.GetTempPath(), "ort_template_smoke_" + Guid.NewGuid().ToString("N"));
            string root = Path.Combine(work, "OUT");
            Directory.CreateDirectory(root);
            try
            {
                // 1. 把样本报告放进"输出根目录"，模拟该机种已有一份报告
                string template = Path.Combine(root, SampleFolder);
                Directory.CreateDirectory(template);
                string sourceOverview = Directory.GetFiles(sample, "*.xls*")
                    .First(f => !Path.GetFileName(f).StartsWith("~$", StringComparison.Ordinal));
                File.Copy(sourceOverview, Path.Combine(template, Path.GetFileName(sourceOverview)));
                byte[] sourceBytes = File.ReadAllBytes(sourceOverview);

                // 2. 已有报告能被认出来（周期号取文件夹名、Cover 与 Waterfall 读得出来）
                Assert.That(ReportTemplateService.FindExistingReportFolder(root, Model), Is.Not.Null,
                    "应能定位该机种已有报告文件夹");
                ExistingReportInfo existing = new ReportTemplateService().FindExistingReport(root, Model);
                Assert.That(existing, Is.Not.Null, "应能读出已有报告信息");
                Assert.That(existing.Week, Is.EqualTo("2616"), "周期号应取自文件夹名");
                Assert.That(existing.Customer, Is.Not.Null.And.Not.Empty, "应读到 Cover 客户别");
                Assert.That(existing.Revision, Is.EqualTo("01P"), "应读到 Cover 版本");
                Assert.That(existing.Stage, Is.EqualTo(PlanStage.MP), "应读到 Cover 阶段");
                Assert.That(existing.SerialNumbers.Count, Is.EqualTo(3), "应读到 3 个序列号");
                Assert.That(existing.WorkOrders.Count, Is.EqualTo(1), "应读到 1 个工令");
                Assert.That(existing.StartDate, Is.Not.Null, "应读到测试起始日期");
                Assert.That(existing.Items.Select(i => i.TestItemName), Is.EquivalentTo(Plan.Select(p => p.Name)),
                    "应从 Waterfall 读出 3 项测试安排");

                // 3. 排期：热冲击 24h=1 天（10/12）、EMI 24h=1 天（10/13）、
                //    Burn In 168h 跨 8 天（10/14〜10/21）——"超过 24 小时按跨天算"
                DateTime start = new(2026, 10, 12);
                List<ReportTemplateItem> items = [];
                foreach ((string name, string category, string hours) in Plan)
                {
                    items.Add(new ReportTemplateItem
                    {
                        TestItemName = name,
                        Category = category,
                        PeriodHours = hours,
                        TestCondition = name + ": smoke"
                    });
                }
                ReportTemplateService.Schedule(items, start);
                Assert.That(items[0].Start, Is.EqualTo(start), "第 1 项从开始日期起排");
                Assert.That(SpanDays(items[0]), Is.EqualTo(1), "24 小时占 1 天");
                Assert.That(items[1].Start, Is.EqualTo(start.AddDays(1)), "第 2 项接在第 1 项之后");
                Assert.That(SpanDays(items[1]), Is.EqualTo(1), "24 小时占 1 天");
                Assert.That(items[2].Start, Is.EqualTo(start.AddDays(2)), "第 3 项接在第 2 项之后");
                Assert.That(SpanDays(items[2]), Is.EqualTo(8), "168 小时应横跨 8 天");
                Assert.That(items[2].End, Is.EqualTo(start.AddDays(9)), "Burn In 结束日 = 起始日 + 7");

                // 4. 生成：概览以已有报告为骨架，产物落在临时目录
                ReportTemplateResult result = new ReportTemplateService().Generate(new ReportTemplateRequest
                {
                    ModelName = Model,
                    Customer = "LENOVO",
                    Revision = "02A",
                    Stage = PlanStage.MP,
                    Period = "WK2633",
                    JobNo = "RT261201",
                    WorkOrder = "WK2601THR331",
                    SerialNumbers = ["SN0001", "SN0002"],
                    StartDate = start,
                    Items = items,
                    OutputRoot = root,
                    CreatedBy = "smoke"
                });
                Assert.That(result.Ok, Is.True, result.Message);

                // 样本报告没被改动，Report 子目录跟着复制过来
                Assert.That(File.ReadAllBytes(sourceOverview), Is.EqualTo(sourceBytes), "不得改动已有报告文件");
                Assert.That(Directory.Exists(Path.Combine(result.Folder, "Report")), Is.True, "应复制 Report 子目录");

                using IWorkbook wb = ExcelNpoi.OpenAny(result.OverviewFile);
                AssertSheetOrder(wb);

                // 5. Cover：客户/版本/阶段填本次的值；测试日期公式保留
                ISheet cover = FindSheet(wb, "cover");
                int coverCol = FindLabelValueColumn(cover, "Customer");
                Assert.That(ExcelNpoi.CellText(cover, 12, coverCol), Is.EqualTo("LENOVO"), "Cover 客户别");
                Assert.That(ExcelNpoi.CellText(cover, 14, coverCol), Is.EqualTo(Model), "Cover 机种");
                Assert.That(ExcelNpoi.CellText(cover, 16, coverCol), Is.EqualTo("02A"), "Cover 版本");
                Assert.That(ExcelNpoi.CellText(cover, 18, coverCol), Is.EqualTo(PlanStage.MP), "Cover 产品阶段");
                ICell testDate = ExcelNpoi.ExistingCell(cover, 20, coverCol);
                Assert.That(testDate?.CellType, Is.EqualTo(CellType.Formula), "Cover 测试日期公式应保留");
                Assert.That(testDate.CellFormula, Does.Contain("Waterfall!E4"), "测试日期公式仍引用 Waterfall 起始日");

                // 6. Waterfall：序列号 + 工令 + 排期（Burn In 从第 7 列起合并 8 天）
                ISheet waterfall = FindSheet(wb, "waterfall");
                Assert.That(ExcelNpoi.CellText(waterfall, 7, 4), Is.EqualTo("SN0001"), "第 1 个序列号");
                Assert.That(ExcelNpoi.CellText(waterfall, 8, 4), Is.EqualTo("SN0002"), "第 2 个序列号");
                Assert.That(ExcelNpoi.CellText(waterfall, 9, 4), Is.EqualTo("WK2601THR331"), "工令写在序列号下面一行");
                Assert.That(ExcelNpoi.CellText(waterfall, 7, 5), Is.EqualTo("Thermal Shock Test"), "第 1 项测试");
                Assert.That(ExcelNpoi.CellText(waterfall, 7, 6), Is.EqualTo("Conducted EMI Measurement"), "第 2 项测试");
                Assert.That(ExcelNpoi.CellText(waterfall, 7, 7), Is.EqualTo("Burn In Test"), "第 3 项测试");
                var burnIn = waterfall.MergedRegions.FirstOrDefault(r => r.FirstRow == 6 && r.FirstColumn == 6);
                Assert.That(burnIn, Is.Not.Null, "Burn In 应有合并区");
                Assert.That(burnIn.LastColumn - burnIn.FirstColumn + 1, Is.EqualTo(8), "Burn In 合并区应跨 8 天");

                // 7. TestStatus：每项测试的 total 列（E，UNITS UNDER TEST）置空
                ISheet status = FindSheet(wb, "teststatus");
                int itemRows = 0;
                for (int r = 9; r <= 24; r++)
                {
                    string name = ExcelNpoi.CellText(status, r, 4);
                    if (name.Length == 0 || name.Contains("ENVIRONMENT") || name.Trim().Equals("EMC", StringComparison.OrdinalIgnoreCase))
                    {
                        continue; // 空行 / 分类行
                    }
                    itemRows++;
                    Assert.That(ExcelNpoi.CellText(status, r, 5), Is.EqualTo(""), $"测试项「{name.Trim()}」的 total 列应置空");
                }
                Assert.That(itemRows, Is.EqualTo(Plan.Length), "应重建出计划里的每一项测试");
            }
            finally
            {
                try
                {
                    Directory.Delete(work, true);
                }
                catch (Exception)
                {
                    // 清理失败不影响结论
                }
            }
        }

        /// <summary>某项测试占几天（含首尾）</summary>
        private static int SpanDays(ReportTemplateItem item)
            => (int)(item.End.Value.Date - item.Start.Value.Date).TotalDays + 1;

        private static ISheet FindSheet(IWorkbook wb, string keyword)
        {
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                if (wb.GetSheetName(i).ToLowerInvariant().Replace(" ", "").Contains(keyword))
                {
                    return wb.GetSheetAt(i);
                }
            }
            Assert.Fail($"找不到工作表：{keyword}");
            return null;
        }

        /// <summary>找"标签所在单元格右边第一个非空单元格"的列号（Cover 的取值列在合并区中间，不能写死列号）</summary>
        private static int FindLabelValueColumn(ISheet sheet, string label)
        {
            for (int r = 1; r <= ExcelNpoi.LastRow(sheet); r++)
            {
                for (int c = 1; c <= ExcelNpoi.LastColumn(sheet); c++)
                {
                    string text = ExcelNpoi.CellText(sheet, r, c);
                    if (text != null && text.StartsWith(label, StringComparison.OrdinalIgnoreCase))
                    {
                        for (int v = c + 1; v <= ExcelNpoi.LastColumn(sheet); v++)
                        {
                            if (!string.IsNullOrWhiteSpace(ExcelNpoi.CellText(sheet, r, v)))
                            {
                                return v;
                            }
                        }
                    }
                }
            }
            Assert.Fail($"找不到标签「{label}」的取值列");
            return 0;
        }

        private static void AssertSheetOrder(IWorkbook wb)
        {
            string[] expected = ["Cover", "ORT Plan", "Waterfall", "TestStatus"];
            List<string> actual = [];
            for (int i = 0; i < wb.NumberOfSheets && i < expected.Length; i++)
            {
                actual.Add(wb.GetSheetName(i));
            }
            Assert.That(actual, Is.EqualTo(expected), "工作表顺序应为 Cover / ORT Plan / Waterfall / TestStatus");
        }
    }
}

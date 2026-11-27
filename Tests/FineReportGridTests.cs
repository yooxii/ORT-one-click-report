using NUnit.Framework;
using ORT一键报告.Models;
using ORT一键报告.Services;
using System.Collections.Generic;

namespace ORT一键报告.Tests
{
    /// <summary>
    /// 帆软取数响应的解析测试：
    /// 1) pageResult 的真实形状是「行数组的数组」（[[单元格…],[单元格…]]），早期按「页 → 行 → 单元格」
    ///    三层解析会一个单元格都取不到（表现为「日志里响应很大、结果显示 0 行」）；
    /// 2) 合并单元格按 rowSpan/columnSpan 铺开；
    /// 3) 分页看板按 totalPage 读页数；
    /// 4) 解析出的网格要能被看板表头识别成数据行。
    /// </summary>
    [TestFixture]
    public class FineReportGridTests
    {
        /// <summary>造一个帆软单元格（value 是「JSON 套 JSON」）</summary>
        private static string Cell(int x, int y, string text, int columnSpan = 1, int rowSpan = 1)
            => "{\"position\":{\"x\":" + x + ",\"y\":" + y + ",\"columnSpan\":" + columnSpan
               + ",\"rowSpan\":" + rowSpan + "},\"value\":"
               + "\"{\\\"type\\\":\\\"text\\\",\\\"value\\\":\\\"" + text + "\\\"}\"}";

        /// <summary>真实形状：pageResult 直接是行数组的数组</summary>
        private static string RowsArrayJson()
            => "{\"pageIndex\":1,\"totalPage\":1,\"pageResult\":["
               + "[" + Cell(0, 0, "依小工令每天刷新") + "," + Cell(1, 0, "") + "],"
               + "[" + Cell(0, 1, "日期") + "," + Cell(1, 1, "類別") + "," + Cell(2, 1, "機種")
               + "," + Cell(3, 1, "生產線別") + "," + Cell(4, 1, "生產日期") + "," + Cell(5, 1, "工令")
               + "," + Cell(6, 1, "生產數量") + "," + Cell(7, 1, "生產時數") + "," + Cell(8, 1, "65報工否")
               + "," + Cell(9, 1, "SIM87入庫否") + "],"
               + "[" + Cell(0, 2, "20261101") + "," + Cell(1, 2, "ORT") + "," + Cell(2, 2, "ADL030-00SG")
               + "," + Cell(3, 2, "P09") + "," + Cell(4, 2, "20261102") + "," + Cell(5, 2, "AH00L30S0037")
               + "," + Cell(6, 2, "2156") + "," + Cell(7, 2, "5.99") + "," + Cell(8, 2, "N")
               + "," + Cell(9, 2, "Y") + "]"
               + "]}";

        [Test]
        public void 行数组的数组_能解析成网格()
        {
            List<List<string>> grid = FineReportClient.ParseGrid(RowsArrayJson());

            Assert.That(grid.Count, Is.EqualTo(3), "标题行 + 表头 + 1 行数据");
            Assert.That(grid[1][0], Is.EqualTo("日期"));
            Assert.That(grid[1][2], Is.EqualTo("機種"));
            Assert.That(grid[2][0], Is.EqualTo("20261101"));
            Assert.That(grid[2][2], Is.EqualTo("ADL030-00SG"));
            Assert.That(grid[2][9], Is.EqualTo("Y"));
        }

        [Test]
        public void 三层形状_页下行下单元格_也能解析()
        {
            string json = "{\"pageIndex\":1,\"pageResult\":[["
                          + "[" + Cell(0, 0, "日期") + "," + Cell(1, 0, "機種") + "],"
                          + "[" + Cell(0, 1, "20261102") + "," + Cell(1, 1, "FSF045-611G") + "]"
                          + "]]}";

            List<List<string>> grid = FineReportClient.ParseGrid(json);

            Assert.That(grid.Count, Is.EqualTo(2));
            Assert.That(grid[0][1], Is.EqualTo("機種"));
            Assert.That(grid[1][1], Is.EqualTo("FSF045-611G"));
        }

        [Test]
        public void 合并单元格_按跨度铺开()
        {
            string json = "{\"pageResult\":[["
                          + Cell(0, 0, "標題", columnSpan: 3)
                          + "],[" + Cell(0, 1, "A") + "]]}";

            List<List<string>> grid = FineReportClient.ParseGrid(json);

            Assert.That(grid[0].Count, Is.EqualTo(3), "列跨度 3 要铺成 3 列");
            Assert.That(grid[0][0], Is.EqualTo("標題"));
            Assert.That(grid[0][2], Is.EqualTo("標題"));
            Assert.That(grid[1][0], Is.EqualTo("A"));
        }

        [Test]
        public void 空响应或错误响应_返回空网格()
        {
            Assert.That(FineReportClient.ParseGrid(null), Is.Empty);
            Assert.That(FineReportClient.ParseGrid(""), Is.Empty);
            Assert.That(FineReportClient.ParseGrid("{\"errorCode\":\"12400007\",\"errorMsg\":\"未登录\"}"), Is.Empty);
            Assert.That(FineReportClient.ParseGrid("not json at all"), Is.Empty);
        }

        [Test]
        public void 末尾整行空_被裁掉()
        {
            string json = "{\"pageResult\":[["
                          + Cell(0, 0, "日期") + "],[" + Cell(0, 1, "") + "]]}";

            List<List<string>> grid = FineReportClient.ParseGrid(json);

            Assert.That(grid.Count, Is.EqualTo(1));
            Assert.That(grid[0][0], Is.EqualTo("日期"));
        }

        [Test]
        public void 总页数_数字与字符串都能认_缺失按一页()
        {
            Assert.That(FineReportClient.ReadTotalPage("{\"totalPage\":3}"), Is.EqualTo(3));
            Assert.That(FineReportClient.ReadTotalPage("{\"totalPage\":\"2\"}"), Is.EqualTo(2));
            Assert.That(FineReportClient.ReadTotalPage("{\"totalPage\":0}"), Is.EqualTo(1));
            Assert.That(FineReportClient.ReadTotalPage("{\"totalPage\":\"\"}"), Is.EqualTo(1));
            Assert.That(FineReportClient.ReadTotalPage("{\"pageIndex\":1}"), Is.EqualTo(1));
            Assert.That(FineReportClient.ReadTotalPage("垃圾"), Is.EqualTo(1));
        }

        [Test]
        public void 请求体_页码可覆盖且参数双重编码()
        {
            OrtBoardEndpoints ep = new()
            {
                DateFromKey = "RUN_DATE",
                DateToKey = "RUN_DATE2",
                MoTypeKey = "MO_TYPE",
                DateLabelKey = "LABEL0_C",
                DateLabelText = "日期:",
                MoTypeLabelKey = "LABEL0_C_C",
                MoTypeLabelText = "类型:",
                PageIndex = 1
            };
            OrtBoardQuery query = new() { DateFrom = "20261101", DateTo = "20261126", MoType = "" };

            string body = FineReportClient.BuildLoadBody(ep, "report2", query, 2);

            Assert.That(body, Does.Contain("pageIndex=2"));
            // __parameters__ 必须是「编码两次」：编码后的 % 再被编码成 %25
            Assert.That(body, Does.Contain("__parameters__=%257B"));
            Assert.That(body, Does.Not.Contain("__parameters__={"));
        }

        [Test]
        public void 网格能被识别成看板数据行()
        {
            List<List<string>> grid = FineReportClient.ParseGrid(RowsArrayJson());

            List<OrtBoardRow> rows = OrtBoardService.ToBoardRows(grid, "tester", out string error);

            Assert.That(error, Is.Null);
            Assert.That(rows.Count, Is.EqualTo(1));
            OrtBoardRow row = rows[0];
            Assert.That(row.Month, Is.EqualTo("202611"));
            Assert.That(row.ReportDate, Is.EqualTo("20261101"));
            Assert.That(row.Category, Is.EqualTo("ORT"));
            Assert.That(row.Model, Is.EqualTo("ADL030-00SG"));
            Assert.That(row.Line, Is.EqualTo("P09"));
            Assert.That(row.ProduceDate, Is.EqualTo("20261102"));
            Assert.That(row.JobNo, Is.EqualTo("AH00L30S0037"));
            Assert.That(row.Qty, Is.EqualTo("2156"));
            Assert.That(row.Hours, Is.EqualTo("5.99"));
            Assert.That(row.Reported65, Is.EqualTo("N"));
            Assert.That(row.Stocked87, Is.EqualTo("Y"));
            Assert.That(row.ImportedBy, Is.EqualTo("tester"));
        }
    }
}

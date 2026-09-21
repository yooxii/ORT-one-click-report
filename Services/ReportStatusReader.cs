using NLog;
using NPOI.SS.UserModel;
using ORT一键报告.Models;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 读取报告概览的 TestStatus 表，判定该报告是否"已完成"。
    /// 判定规则（与用户约定一致）：
    /// - 表头行在前 10 行内，按单元格文本（去空白、忽略大小写）定位 4 列：
    ///   TEST ITEMS / UNITS UNDER TEST / TOTAL / FAIL
    /// - 从表头下一行开始遍历到最后一行：跳过 TEST ITEMS 列为空的行（分类标题行 / 合计行）
    /// - 对每一个测试项行：UNITS UNDER TEST == TOTAL + FAIL → 该行完成
    /// - 所有测试项行都完成 → 已完成；否则 → 进行中（含全部为 0 的未开始情况）
    /// 找不到 TestStatus 表 / 找不到表头 / 一个测试项行都没有 → 返回 null（调用方保持原值不改）
    /// </summary>
    public class ReportStatusReader
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// 读取报告概览文件，返回报告状态；读不到 / 解析失败返回 null
        /// </summary>
        public string ReadStatus(string overviewFile)
        {
            if (string.IsNullOrWhiteSpace(overviewFile) || !File.Exists(overviewFile))
            {
                return null;
            }
            try
            {
                IWorkbook wb = ExcelNpoi.OpenAny(overviewFile);
                try
                {
                    return ReadStatusFromWorkbook(wb);
                }
                finally
                {
                    wb.Close();
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"读取报告状态失败（{overviewFile}）: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 从工作簿里找 TestStatus 表并判定状态
        /// </summary>
        private string ReadStatusFromWorkbook(IWorkbook wb)
        {
            ISheet sheet = FindTestStatusSheet(wb);
            if (sheet == null)
            {
                return null;
            }
            (int headerRow, Dictionary<string, int> columns) = FindHeaderRow(sheet);
            if (headerRow == 0)
            {
                return null;
            }
            int colTestItem = columns["TESTITEMS"];
            int colUnits = columns["UNITSUNDERTEST"];
            int colTotal = columns["TOTAL"];
            int colFail = columns["FAIL"];

            int lastRow = ExcelNpoi.LastRow(sheet);
            int itemCount = 0;
            int completedCount = 0;
            for (int r = headerRow + 1; r <= lastRow; r++)
            {
                string testName = ExcelNpoi.CellText(sheet, r, colTestItem)?.Trim();
                if (string.IsNullOrWhiteSpace(testName))
                {
                    continue; // 分类标题行 / 合计行 / 空行
                }
                // 跳过合计行（TEST ITEMS 列写 Total / 合計 / 合计 等）
                if (IsTotalRowLabel(testName))
                {
                    continue;
                }
                if (!TryReadNumber(sheet, r, colUnits, out double units))
                {
                    // UNITS UNDER TEST 读不到数值：该行视为未完成
                    itemCount++;
                    continue;
                }
                bool totalOk = TryReadNumber(sheet, r, colTotal, out double total);
                bool failOk = TryReadNumber(sheet, r, colFail, out double fail);
                itemCount++;
                if (totalOk && failOk && Math.Abs(units - (total + fail)) < 0.0001)
                {
                    completedCount++;
                }
            }
            if (itemCount == 0)
            {
                return null; // 没有测试项行，判不出结果
            }
            return completedCount == itemCount ? ReportStatusKind.Complete : ReportStatusKind.InProgress;
        }

        /// <summary>
        /// 找 TestStatus 工作表：名字精确匹配（忽略大小写）优先，其次包含 "test status"
        /// </summary>
        private static ISheet FindTestStatusSheet(IWorkbook wb)
        {
            if (wb == null)
            {
                return null;
            }
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                ISheet sheet = wb.GetSheetAt(i);
                if (string.Equals(sheet.SheetName, "TestStatus", StringComparison.OrdinalIgnoreCase))
                {
                    return sheet;
                }
            }
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                ISheet sheet = wb.GetSheetAt(i);
                string normalized = Normalize(sheet.SheetName);
                if (normalized.Contains("TESTSTATUS"))
                {
                    return sheet;
                }
            }
            return null;
        }

        /// <summary>
        /// 在前 10 行里找同时包含 4 个关键列的表头行；返回 (行号, 规范化列名→列号) 映射
        /// </summary>
        private static (int, Dictionary<string, int>) FindHeaderRow(ISheet sheet)
        {
            int endRow = Math.Min(ExcelNpoi.LastRow(sheet), 10);
            int endCol = ExcelNpoi.LastColumn(sheet);
            string[] required = ["TESTITEMS", "UNITSUNDERTEST", "TOTAL", "FAIL"];
            for (int r = 1; r <= endRow; r++)
            {
                Dictionary<string, int> map = new();
                for (int c = 1; c <= endCol; c++)
                {
                    string key = Normalize(ExcelNpoi.CellText(sheet, r, c));
                    if (key.Length == 0 || map.ContainsKey(key))
                    {
                        continue;
                    }
                    map[key] = c;
                }
                bool all = true;
                foreach (string need in required)
                {
                    if (!ContainsKey(map, need))
                    {
                        all = false;
                        break;
                    }
                }
                if (all)
                {
                    // 把命中的列名统一映射回 required 里的规范键，方便上层按固定键取列号
                    Dictionary<string, int> result = new();
                    foreach (string need in required)
                    {
                        result[need] = FindColumn(map, need);
                    }
                    return (r, result);
                }
            }
            return (0, null);
        }

        /// <summary>表头映射里是否存在"包含指定关键字"的列（关键字本身已规范化）</summary>
        private static bool ContainsKey(Dictionary<string, int> map, string normalizedNeed)
            => FindColumn(map, normalizedNeed) > 0;

        /// <summary>在表头映射里找"包含指定关键字"的第一列列号；没有返回 0</summary>
        private static int FindColumn(Dictionary<string, int> map, string normalizedNeed)
        {
            foreach (KeyValuePair<string, int> kv in map)
            {
                if (kv.Key.Contains(normalizedNeed))
                {
                    return kv.Value;
                }
            }
            return 0;
        }

        /// <summary>规范化表头文本：去掉所有空白、转大写（"UNITS UNDER TEST" → "UNITSUNDERTEST"）</summary>
        private static string Normalize(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }
            System.Text.StringBuilder sb = new(text.Length);
            foreach (char ch in text)
            {
                if (!char.IsWhiteSpace(ch))
                {
                    sb.Append(char.ToUpperInvariant(ch));
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// TEST ITEMS 列的文本是否是"合计行"标签（Total / 合計 / 合计 / 總計 / 总计）
        /// </summary>
        private static bool IsTotalRowLabel(string testName)
        {
            string normalized = Normalize(testName);
            return normalized == "TOTAL"
                || normalized.Contains("合計")
                || normalized.Contains("合计")
                || normalized.Contains("總計")
                || normalized.Contains("总计");
        }

        /// <summary>
        /// 读单元格的数值：支持数字格式与文本格式的数字；错误值（#N/A 等）/ 空 / 非数字 → false
        /// </summary>
        private static bool TryReadNumber(ISheet sheet, int row, int col, out double value)
        {
            value = 0;
            ICell cell = ExcelNpoi.ExistingCell(sheet, row, col);
            if (cell == null)
            {
                return false;
            }
            try
            {
                if (cell.CellType == CellType.Numeric)
                {
                    value = cell.NumericCellValue;
                    return true;
                }
                if (cell.CellType == CellType.Formula)
                {
                    if (cell.CachedFormulaResultType == CellType.Numeric)
                    {
                        value = cell.NumericCellValue;
                        return true;
                    }
                    return false; // 公式结果是错误值 / 字符串 → 视为不可用
                }
                string text = ExcelNpoi.CellText(sheet, row, col)?.Trim();
                if (string.IsNullOrWhiteSpace(text))
                {
                    return false;
                }
                return double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out value);
            }
            catch
            {
                return false;
            }
        }
    }
}

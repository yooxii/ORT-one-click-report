using NPOI.SS.UserModel;
using ORT一键报告.Models;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace ORT一键报告.Utils
{
    /// <summary>ATE 原始数据里的一行：一个样品的全部测试值</summary>
    public class AteRow
    {
        public string SN { get; set; } = "";
        public List<string> Values { get; set; } = [];
    }

    /// <summary>
    /// ATE 原始数据表：测试条件（S/N 上一行）、测试项目名（S/N 行）、规格上下限、所有数据行。
    /// 原始文件里数据从 "S/N" 的下一行开始，到 "MAX_SPEC" 的上一行结束，
    /// "MAX_SPEC" / "MIN_SPEC" 两行分别是每个项目的上下限。
    /// </summary>
    public class AteSheet
    {
        public List<string> Conditions { get; set; } = [];
        public List<string> OutputTypes { get; set; } = [];
        public List<string> MaxSpecs { get; set; } = [];
        public List<string> MinSpecs { get; set; } = [];
        public List<AteRow> Rows { get; set; } = [];

        public int ItemCount => OutputTypes.Count;

        /// <summary>
        /// 试验前/试验后各最多能选的数据条数：原始数据条数的一半（数据一般是成对的"前/后"）。
        /// 条数为奇数时无法按"前/后"配对，不做上限限制。
        /// </summary>
        public int MaxPerGroup => Rows.Count % 2 == 0 ? Rows.Count / 2 : Rows.Count;
    }

    /// <summary>
    /// ATE 数据的读取/分组猜测/上下限判定，以及按 ATE 报告模板生成报告文件。
    /// 从 ATEWindow 里抽出来，便于单独验证（不含任何界面依赖）。
    /// </summary>
    public static class AteReportData
    {
        /* ###############################  读取  ################################ */

        /// <summary>
        /// 读取 ATE 原始数据（.xls / .xlsx 都支持）；读不到 "S/N"/"MAX_SPEC" 时返回 null
        /// </summary>
        public static AteSheet Read(string fileName)
        {
            IWorkbook wb = ExcelNpoi.OpenAny(fileName);
            try
            {
                ISheet ws = ExcelNpoi.SheetAt(wb, 0);
                DataCell startCell = Report.FindCellByValue(ws, "s/n");
                DataCell maxSpecCell = Report.FindCellByValue(ws, "MAX_SPEC");
                if (startCell == null || maxSpecCell == null || startCell.Row >= maxSpecCell.Row)
                {
                    return null;
                }
                AteSheet sheet = new();
                int lastCol = ExcelNpoi.LastColumn(ws);
                for (int c = startCell.Column + 1; c <= lastCol; c++)
                {
                    if (string.IsNullOrWhiteSpace(ExcelNpoi.CellText(ws, startCell.Row, c)))
                    {
                        break;
                    }
                    sheet.Conditions.Add(ExcelNpoi.CellText(ws, startCell.Row - 1, c));
                    sheet.OutputTypes.Add(ExcelNpoi.CellText(ws, startCell.Row, c));
                    sheet.MaxSpecs.Add(ExcelNpoi.CellText(ws, maxSpecCell.Row, c));
                    sheet.MinSpecs.Add(ExcelNpoi.CellText(ws, maxSpecCell.Row + 1, c));
                }
                if (sheet.ItemCount == 0)
                {
                    return null;
                }
                for (int r = startCell.Row + 1; r < maxSpecCell.Row; r++)
                {
                    string sn = ExcelNpoi.CellText(ws, r, startCell.Column);
                    AteRow row = new() { SN = sn };
                    bool any = !string.IsNullOrWhiteSpace(sn);
                    for (int i = 0; i < sheet.ItemCount; i++)
                    {
                        string text = ExcelNpoi.CellText(ws, r, startCell.Column + 1 + i);
                        any |= !string.IsNullOrWhiteSpace(text);
                        row.Values.Add(text);
                    }
                    if (any)
                    {
                        sheet.Rows.Add(row);
                    }
                }
                return sheet.Rows.Count == 0 ? null : sheet;
            }
            finally
            {
                wb.Close();
            }
        }

        /// <summary>
        /// 猜测"试验前/试验后"的分组（旧逻辑：相邻成对 → 前半/后半）。
        /// 判断不出来时返回 null，由使用者手工选择（这部分数据是人工填写的，规律不可靠）。
        /// </summary>
        public static bool[] GuessIsBefore(IList<string> sns)
        {
            if (sns == null || sns.Count < 2)
            {
                return null;
            }
            int half = sns.Count / 2;
            int flag;
            if (IsPair(sns[0], sns[1]))
            {
                flag = 1;
            }
            else if (half > 0 && IsPair(sns[0], sns[half]))
            {
                flag = 2;
            }
            else
            {
                return null;   // 判断不出来
            }
            bool[] result = new bool[sns.Count];
            for (int i = 0; i < sns.Count; i++)
            {
                result[i] = flag == 1
                    ? (i + 1) % 2 == 1      // 相邻成对：奇数为试验前
                    : i < half;             // 前半试验前、后半试验后
            }
            return result;
        }

        private static bool IsPair(string a, string b, char aC = '1', char bC = '2')
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b) || a.Length != b.Length)
            {
                return false;
            }
            int diffCount = 0;
            int diffIndex = -1;
            for (int i = 0; i < a.Length; i++)
            {
                if (a[i] != b[i])
                {
                    diffCount++;
                    diffIndex = i;
                    if (diffCount > 1)
                    {
                        return false;
                    }
                }
            }
            if (diffCount != 1)
            {
                return false;
            }
            char c1 = a[diffIndex];
            char c2 = b[diffIndex];
            return (c1 == bC && c2 == aC) || (c1 == aC && c2 == bC);
        }

        /* ###############################  上下限  ################################ */

        /// <summary>数值文本解析；空、"*" 或非数值 → false（表示这一侧不限制）</summary>
        public static bool TryParseNumber(string text, out double value)
        {
            value = 0;
            string t = text?.Trim();
            if (string.IsNullOrEmpty(t) || t == "*")
            {
                return false;
            }
            return double.TryParse(t, NumberStyles.Any, CultureInfo.InvariantCulture, out value)
                || double.TryParse(t, NumberStyles.Any, CultureInfo.CurrentCulture, out value);
        }

        /// <summary>
        /// 是否超出上下限：上下限里的空和 "*" 都表示不考虑这一侧；
        /// 上下限都有效但填反了（如 MAX=-40、MIN=-70）时按数值大小纠正。
        /// </summary>
        public static bool IsOutOfSpec(string value, string maxSpec, string minSpec)
        {
            if (!TryParseNumber(value, out double v))
            {
                return false;
            }
            bool hasMax = TryParseNumber(maxSpec, out double max);
            bool hasMin = TryParseNumber(minSpec, out double min);
            if (hasMax && hasMin && max < min)
            {
                (max, min) = (min, max);
            }
            if (hasMax && v > max)
            {
                return true;
            }
            return hasMin && v < min;
        }

        /// <summary>限值文本（用于界面显示）：空/"*" 显示为 "-"</summary>
        public static string SpecText(string spec) => string.IsNullOrWhiteSpace(spec) ? "-" : spec.Trim();

        /// <summary>列标题：测试项目名（如 Vout_1_3）</summary>
        public static string ColumnTitle(AteSheet sheet, int index)
            => index >= 0 && index < sheet.ItemCount ? sheet.OutputTypes[index] : "";

        /* ###############################  写出报告  ################################ */

        // ATE 报告模板的固定结构（1 基行号）：表头/项目名/规格/数据起始行，以及试验前、试验后各 3 行
        private const int ItemNameRow = 3;
        private const int OutputTypeRow = 4;
        private const int FirstDataRow = 5;
        private const int TemplateGroupRows = 3;
        private const int MaxSpecRow = 11;
        private const int MinSpecRow = 12;
        private const int FirstColumn = 4;   // 第一个项目的列（D）
        private const int SnColumn = 3;      // S/N 列（C）
        private const int PreStatsRow = 13;  // 试验前统计块（Min/Max/Average/Judgement，4 行）
        private const int PostStatsRow = 17; // 试验后统计块

        /// <summary>
        /// 按模板生成 ATE 报告：试验前数据在前、试验后数据在后，行数随选择的数量增减。
        /// </summary>
        public static void Write(AteSheet sheet, IList<AteRow> beforeRows, IList<AteRow> afterRows, string templatePath, string outputPath)
        {
            IWorkbook wb = ExcelNpoi.OpenAny(templatePath);
            try
            {
                ISheet ws = ExcelNpoi.SheetAt(wb, 0);
                int lastRow = ExcelNpoi.LastRow(ws);

                // 1. 表头（测试条件）、项目名、规格上下限，并复制首列样式/公式
                for (int i = 0; i < sheet.ItemCount; i++)
                {
                    int col = FirstColumn + i;
                    ExcelNpoi.SetCell(ws, ItemNameRow, col, sheet.Conditions[i]);
                    ExcelNpoi.SetCell(ws, OutputTypeRow, col, sheet.OutputTypes[i]);
                    SetCellAuto(ws, MaxSpecRow, col, sheet.MaxSpecs[i]);
                    SetCellAuto(ws, MinSpecRow, col, sheet.MinSpecs[i]);
                    if (i > 0)
                    {
                        ExcelNpoi.CopyColumnStylesAndFormulas(ws, FirstColumn, col, 1, lastRow);
                    }
                }

                // 2. 按选择的数量增减行（试验前块起点固定，试验后块随之平移）
                int beforeCount = beforeRows?.Count ?? 0;
                int afterCount = afterRows?.Count ?? 0;
                EditGroupRows(ws, FirstDataRow, beforeCount);
                EditGroupRows(ws, FirstDataRow + beforeCount, afterCount);
                // 两组行数变化量之和：规格行与统计块整体平移这么多行
                int delta = (beforeCount - TemplateGroupRows) + (afterCount - TemplateGroupRows);

                // 3. 某一侧没有数据时，把该侧的统计块整块删掉（避免统计到空区域/错位的行）。
                //    先删靠后的试验后统计块，再删试验前统计块，位置才不会互相影响。
                if (afterCount == 0)
                {
                    ExcelNpoi.DeleteRows(ws, PostStatsRow + delta, 4);
                }
                if (beforeCount == 0)
                {
                    ExcelNpoi.DeleteRows(ws, PreStatsRow + delta, 4);
                }

                // 4. 写数据：试验前若干行 + 试验后若干行
                int row = FirstDataRow;
                WriteRows(ws, beforeRows, ref row);
                WriteRows(ws, afterRows, ref row);

                // 5. 重写统计/判定公式（行数变了，模板里写死的区域要跟着改），
                //    并直接写入算好的结果，保证不打开 Excel 也能看到正确的统计值。
                int specMaxRow = FirstDataRow + beforeCount + afterCount;
                WriteStats(ws, sheet, beforeRows, afterRows, specMaxRow, specMaxRow + 1);

                ExcelNpoi.Save(wb, outputPath);
            }
            finally
            {
                wb.Close();
            }
        }

        /// <summary>
        /// 写试验前/试验后的统计块（Min / Max / Average / 判定，每块 4 行）。
        /// 公式与模板一致：判定 = 最大值不超过 MAX_SPEC 且最小值不低于 MIN_SPEC
        /// （上下限为空或 "*" 表示该侧不限制）。
        /// </summary>
        private static void WriteStats(ISheet ws, AteSheet sheet, IList<AteRow> beforeRows, IList<AteRow> afterRows, int specMaxRow, int specMinRow)
        {
            int cursor = specMinRow + 1;
            List<AteRow> before = [.. (beforeRows ?? [])];
            List<AteRow> after = [.. (afterRows ?? [])];
            if (before.Count > 0)
            {
                WritePhaseStats(ws, sheet, before, cursor, specMaxRow, specMinRow, FirstDataRow, FirstDataRow + before.Count - 1);
                cursor += 4;
            }
            if (after.Count > 0)
            {
                WritePhaseStats(ws, sheet, after, cursor, specMaxRow, specMinRow, FirstDataRow + before.Count, FirstDataRow + before.Count + after.Count - 1);
            }
        }

        private static void WritePhaseStats(ISheet ws, AteSheet sheet, List<AteRow> rows, int firstRow, int specMaxRow, int specMinRow, int dataFirstRow, int dataLastRow)
        {
            int minRow = firstRow;
            int maxRow = firstRow + 1;
            int avgRow = firstRow + 2;
            int judgeRow = firstRow + 3;
            for (int i = 0; i < sheet.ItemCount; i++)
            {
                int col = FirstColumn + i;
                string letter = ColumnLetter(col);
                string range = $"{letter}{dataFirstRow}:{letter}{dataLastRow}";
                double min = double.MaxValue;
                double max = double.MinValue;
                double sum = 0;
                int count = 0;
                bool outOfSpec = false;
                foreach (AteRow data in rows)
                {
                    string text = i < data.Values.Count ? data.Values[i] : "";
                    if (!TryParseNumber(text, out double value))
                    {
                        continue;
                    }
                    min = Math.Min(min, value);
                    max = Math.Max(max, value);
                    sum += value;
                    count++;
                    outOfSpec |= IsOutOfSpec(text, sheet.MaxSpecs[i], sheet.MinSpecs[i]);
                }
                string maxRef = $"{letter}${specMaxRow}";
                string minRef = $"{letter}${specMinRow}";
                SetFormulaWithValue(ws, minRow, col, $"MIN({range})", count > 0 ? min : 0);
                SetFormulaWithValue(ws, maxRow, col, $"MAX({range})", count > 0 ? max : 0);
                SetFormulaWithValue(ws, avgRow, col, $"AVERAGE({range})", count > 0 ? sum / count : 0);
                ExcelNpoi.SetFormula(ws, judgeRow, col,
                    $"IF(AND(OR({maxRef}=\"*\",{maxRef}=\"\",{letter}{maxRow}<=MAX({maxRef},{minRef})),"
                    + $"OR({minRef}=\"*\",{minRef}=\"\",{letter}{minRow}>=MIN({maxRef},{minRef}))),\"PASS\",\"FAIL\")");
                ExcelNpoi.SetCell(ws, judgeRow, col, outOfSpec ? "FAIL" : "PASS");
            }
        }

        /// <summary>
        /// 写公式并同时写入缓存结果：这样 Excel 打开时公式是活的（改数据会自动重算），
        /// 用别的工具/不重算时看到的也是正确数值。
        /// </summary>
        private static void SetFormulaWithValue(ISheet ws, int row1, int col1, string formula, double value)
        {
            ExcelNpoi.SetFormula(ws, row1, col1, formula);
            ExcelNpoi.SetCell(ws, row1, col1, value);
        }

        /// <summary>1 基列号 → 列字母（4 → "D"）</summary>
        private static string ColumnLetter(int col1)
            => ExcelNpoi.AddressOf(1, col1).TrimEnd('0', '1', '2', '3', '4', '5', '6', '7', '8', '9');

        private static void WriteRows(ISheet ws, IList<AteRow> rows, ref int row)
        {
            foreach (AteRow data in rows ?? [])
            {
                ExcelNpoi.SetCell(ws, row, SnColumn, data.SN);
                for (int i = 0; i < data.Values.Count; i++)
                {
                    SetCellAuto(ws, row, FirstColumn + i, data.Values[i]);
                }
                row++;
            }
        }

        /// <summary>
        /// 能解析成数字的文本按数字写、其余按文本写。
        /// 模板里的 MIN/MAX/AVERAGE/判定 公式只统计数字单元格，写成文本会让统计全部变成 0
        /// （"*"、空值等限制标记仍按文本写）。
        /// </summary>
        private static void SetCellAuto(ISheet ws, int row, int col, string text)
        {
            if (TryParseNumber(text, out double value))
            {
                ExcelNpoi.SetCell(ws, row, col, value);
            }
            else
            {
                ExcelNpoi.SetCell(ws, row, col, text ?? "");
            }
        }

        /// <summary>
        /// 把某个分组的数据行数调整成 count：模板里每组固定 3 行，多则插入、少则删除（并复制行样式）
        /// </summary>
        private static void EditGroupRows(ISheet ws, int firstRow, int count)
        {
            if (count == TemplateGroupRows)
            {
                return;
            }
            int columns = ExcelNpoi.LastColumn(ws);
            if (count > TemplateGroupRows)
            {
                int n = count - TemplateGroupRows;
                ExcelNpoi.InsertRows(ws, firstRow + 1, n);
                for (int i = 0; i < n; i++)
                {
                    ExcelNpoi.CopyRowStyle(ws, firstRow, firstRow + 1 + i, columns);
                }
            }
            else
            {
                ExcelNpoi.DeleteRows(ws, firstRow, TemplateGroupRows - count);
            }
        }
    }
}

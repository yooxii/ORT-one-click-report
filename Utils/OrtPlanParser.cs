using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ORT一键报告.Utils
{
    /// <summary>
    /// ORT Plan 表的一行（分类行或测试项行）
    /// </summary>
    public class ParsedOrtPlanRow
    {
        /// <summary>是否为分类行（如 RELIABILITY TEST / EMC）</summary>
        public bool IsCategory { get; set; }

        /// <summary>所属分类（分类行即自身文本）</summary>
        public string Category { get; set; }

        /// <summary>测试项目名（分类行为空）</summary>
        public string TestItemName { get; set; }

        /// <summary>抽样计划</summary>
        public string SamplingPlan { get; set; }

        /// <summary>测试条件</summary>
        public string TestCondition { get; set; }

        /// <summary>通过判定</summary>
        public string PassCriterion { get; set; }

        /// <summary>备注</summary>
        public string Remark { get; set; }
    }

    /// <summary>
    /// 一份报告 ORT Plan 表的解析结果
    /// </summary>
    public class ParsedOrtPlan
    {
        /// <summary>按原表顺序的行（含分类行）</summary>
        public List<ParsedOrtPlanRow> Rows { get; } = [];

        /// <summary>表末 Note 说明（Sampling principle 等）</summary>
        public string Note { get; set; }

        /// <summary>表头所在行（1 基；0 表示没找到，按默认版式解析）</summary>
        public int HeaderRow { get; set; }

        /// <summary>测试项行（不含分类行）</summary>
        public List<ParsedOrtPlanRow> Items
        {
            get
            {
                List<ParsedOrtPlanRow> items = [];
                foreach (ParsedOrtPlanRow row in Rows)
                {
                    if (!row.IsCategory)
                    {
                        items.Add(row);
                    }
                }
                return items;
            }
        }
    }

    /// <summary>
    /// ORT Plan 表解析：各机种报告的"Ongoing Reliability Test Plan"版式基本一致
    /// （表头 NO. / TEST ITEMS / SAMPLING PLAN / TEST CONDITOIN / PASS CRITERION / COMMENT，
    /// 中间夹 RELIABILITY TEST、EMC 等分类行，表末一行 Note 说明），
    /// 这里按关键字定位列，兼容列位置与表头拼写差异（TEST CONDITOIN 是原表里的既有笔误）。
    /// </summary>
    public static class OrtPlanParser
    {
        /// <summary>解析工作簿里的 ORT Plan 表（按名称找，找不到返回 null）</summary>
        public static ParsedOrtPlan ParseWorkbook(IWorkbook workbook)
        {
            if (workbook == null)
            {
                return null;
            }
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                string name = workbook.GetSheetName(i);
                if (!string.IsNullOrWhiteSpace(name) && NormalizeName(name).Contains("ORTPLAN"))
                {
                    return Parse(workbook.GetSheetAt(i));
                }
            }
            return null;
        }

        /// <summary>解析 ORT Plan 工作表</summary>
        public static ParsedOrtPlan Parse(ISheet sheet)
        {
            ParsedOrtPlan plan = new();
            if (sheet == null)
            {
                return plan;
            }
            int lastRow = ExcelNpoi.LastRow(sheet);
            int lastCol = ExcelNpoi.LastColumn(sheet);
            if (lastRow <= 0 || lastCol <= 0)
            {
                return plan;
            }

            // 1. 先定位表头行（该行必然有"抽样计划"的英文 SAMPLING PLAN），再在该行里解析各列
            int headerRow = 0;
            for (int r = 1; r <= lastRow && headerRow == 0; r++)
            {
                for (int c = 1; c <= lastCol; c++)
                {
                    if (NormalizeName(ExcelNpoi.CellText(sheet, r, c)).Contains("SAMPLING"))
                    {
                        headerRow = r;
                        break;
                    }
                }
            }
            int cItem = 0, cSample = 0, cCondition = 0, cCriterion = 0, cRemark = 0;
            if (headerRow > 0)
            {
                for (int c = 1; c <= lastCol; c++)
                {
                    string text = NormalizeName(ExcelNpoi.CellText(sheet, headerRow, c));
                    if (text.Length == 0)
                    {
                        continue;
                    }
                    if (cSample == 0 && text.Contains("SAMPLING"))
                    {
                        cSample = c;
                    }
                    else if (cItem == 0 && text.Contains("TEST") && text.Contains("ITEM"))
                    {
                        cItem = c;
                    }
                    else if (cCondition == 0 && text.Contains("CONDIT"))
                    {
                        cCondition = c;
                    }
                    else if (cCriterion == 0 && text.Contains("CRITERI"))
                    {
                        cCriterion = c;
                    }
                    else if (cRemark == 0 && text.Contains("COMMENT"))
                    {
                        cRemark = c;
                    }
                }
            }
            if (headerRow == 0)
            {
                // 没找到表头：按最常见版式取默认列（B=NO. C=TEST ITEMS D=SAMPLING E=CONDITION F=CRITERION G=COMMENT）
                headerRow = 4;
                cItem = 3;
                cSample = 4;
                cCondition = 5;
                cCriterion = 6;
                cRemark = 7;
            }
            else
            {
                if (cItem == 0) cItem = Math.Max(1, cSample - 1);
                if (cCondition == 0 && cCriterion > 0) cCondition = cCriterion - 1;
                if (cCriterion == 0 && cRemark > 0) cCriterion = cRemark - 1;
                if (cRemark == 0) cRemark = Math.Min(lastCol, cCriterion + 1);
            }
            plan.HeaderRow = headerRow;

            // 2. 逐行解析：空行跳过（同一份计划里也可能夹空行），分类行只换分类，Note 行收说明
            string category = null;
            for (int r = headerRow + 1; r <= lastRow; r++)
            {
                string item = Clean(ExcelNpoi.CellText(sheet, r, cItem));
                string sample = Clean(ExcelNpoi.CellText(sheet, r, cSample));
                string condition = Clean(ExcelNpoi.CellText(sheet, r, cCondition));
                string criterion = Clean(ExcelNpoi.CellText(sheet, r, cCriterion));
                string remark = Clean(ExcelNpoi.CellText(sheet, r, cRemark));

                // 表末 Note 说明：整行通常只有一段文字（落在测试项目列，少数版本落在抽样计划列），
                // 没有其它字段时按说明处理，不当作测试项
                bool onlyItem = sample.Length == 0 && condition.Length == 0 && criterion.Length == 0 && remark.Length == 0;
                string noteCandidate = item.Length > 0 ? item : sample;
                if ((item.Length == 0 || onlyItem) && noteCandidate.StartsWith("Note", StringComparison.OrdinalIgnoreCase))
                {
                    if (string.IsNullOrEmpty(plan.Note))
                    {
                        plan.Note = StripNotePrefix(noteCandidate);
                    }
                    continue;
                }

                if (item.Length == 0)
                {
                    continue;
                }

                if (onlyItem && LooksLikeCategory(item))
                {
                    category = item;
                    plan.Rows.Add(new ParsedOrtPlanRow { IsCategory = true, Category = category, TestItemName = item });
                    continue;
                }

                plan.Rows.Add(new ParsedOrtPlanRow
                {
                    IsCategory = false,
                    Category = category,
                    TestItemName = item,
                    SamplingPlan = sample,
                    TestCondition = condition,
                    PassCriterion = criterion,
                    Remark = remark
                });
            }
            return plan;
        }

        /// <summary>
        /// 文本归一化（用于比较是否"同一段话"）：
        /// 统一换行、去掉不换行空格、℃ 统一成 °C、行首尾与行内连续空白压成一个空格、去掉空行。
        /// </summary>
        public static string Normalize(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return "";
            }
            string value = text.Replace("\r\n", "\n").Replace('\r', '\n').Replace('\u00A0', ' ').Replace("℃", "°C");
            List<string> lines = [];
            foreach (string raw in value.Split('\n'))
            {
                string line = Regex.Replace(raw.Trim(), @"[ \t]+", " ");
                if (line.Length > 0)
                {
                    lines.Add(line);
                }
            }
            return string.Join("\n", lines);
        }

        /// <summary>把文本压成单行（用于标题/名称的关键字匹配）</summary>
        public static string NormalizeName(string text)
            => string.IsNullOrEmpty(text) ? "" : Regex.Replace(text, @"\s+", "").ToUpperInvariant();

        /// <summary>单元格文本清理：统一换行、去掉行尾空白（保留行内换行）</summary>
        private static string Clean(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return "";
            }
            string value = text.Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n\n", "\n");
            List<string> lines = [];
            foreach (string raw in value.Split('\n'))
            {
                string line = raw.TrimEnd();
                if (line.Trim().Length > 0)
                {
                    lines.Add(line);
                }
            }
            return string.Join("\n", lines).Trim();
        }

        /// <summary>去掉 Note 行开头的 "Note." 前缀</summary>
        private static string StripNotePrefix(string text)
            => Regex.Replace(text, @"^\s*Note\s*[.:：]?\s*", "", RegexOptions.IgnoreCase).Trim();

        /// <summary>
        /// 是否像分类名：整段没有小写字母（RELIABILITY TEST / EMC），
        /// 或命中已知分类名。测试项名（Thermal Shock Test）含小写，不会被误判。
        /// </summary>
        private static bool LooksLikeCategory(string text)
        {
            if (text.Length == 0)
            {
                return false;
            }
            string normalized = NormalizeName(text);
            if (normalized is "RELIABILITYTEST" or "EMC" or "ENVIRONMENTTESTS" or "ENVIRONMENTTEST"
                or "SAFETYTEST" or "MECHANICALTEST" or "OTHER")
            {
                return true;
            }
            foreach (char ch in text)
            {
                if (char.IsLower(ch))
                {
                    return false;
                }
            }
            // 全是数字/符号的行（NO. 之类的残留）不算分类
            return Regex.IsMatch(text, "[A-Z]");
        }
    }
}

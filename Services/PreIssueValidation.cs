using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 预领导入的「先看清楚再动手」校验：
    /// 导入前按表头形状判断文件到底是哪一类（月排程 / ORT看板 / 开单数据），
    /// 不是本页该导的数据就直接拦下来并给出该去哪儿导的提示，避免把开单数据当月排程导进去。
    /// </summary>
    public static class PreIssueValidation
    {
        /// <summary>开单/领用数据的典型表头</summary>
        private static readonly string[] VoucherHeaders =
        [
            "VCHNO", "VCHNO_CODE", "PART_NO", "REQ_QTY", "VCH_STATUS", "CLOSE_CODE", "WARE_CODE", "TR_CODE"
        ];

        /// <summary>ORT 看板的典型表头</summary>
        private static readonly string[] BoardHeaders =
        [
            "SIM87入庫否", "SIM87入库否", "SIM94禁止出貨", "SIM94禁止出货", "收料日期", "生產線別", "生产线别"
        ];

        /// <summary>月排程的典型表头（前两列）</summary>
        private static readonly string[] MonthHeaderWords = ["年月", "月份", "年/月", "年 月", "年月日"];

        /// <summary>是否看起来像开单/领用数据</summary>
        public static bool LooksLikeVoucherExport(IEnumerable<string> headers)
        {
            HashSet<string> set = Normalize(headers);
            return VoucherHeaders.Any(h => set.Contains(h.ToUpperInvariant()));
        }

        /// <summary>是否看起来像 ORT 看板数据</summary>
        public static bool LooksLikeOrtBoard(IEnumerable<string> headers)
        {
            HashSet<string> set = Normalize(headers);
            return BoardHeaders.Any(h => set.Contains(h.ToUpperInvariant()));
        }

        /// <summary>是否看起来像月排程：第 1 列是年月、第 2 列是線別，并且有日列</summary>
        public static bool LooksLikeSchedule(IEnumerable<string> headers, out string why)
        {
            List<string> list = (headers ?? []).Select(h => (h ?? "").Trim()).ToList();
            string first = list.Count > 0 ? list[0] : "";
            string second = list.Count > 1 ? list[1] : "";
            int dayColumns = list.Count(h => IsDayHeader(h, out _));
            if (!IsMonthHeader(first))
            {
                why = $"第 1 列应是「年月」，当前是「{first}」";
                return false;
            }
            if (!IsLineHeader(second))
            {
                why = $"第 2 列应是「線別」，当前是「{second}」";
                return false;
            }
            if (dayColumns < 3)
            {
                why = $"没找到「1日 / 2日 …」这样的日列（只认出 {dayColumns} 个）";
                return false;
            }
            why = null;
            return true;
        }

        /// <summary>第 1 列的表头是不是年月类</summary>
        public static bool IsMonthHeader(string text)
        {
            string value = (text ?? "").Trim();
            if (value.Length == 0)
            {
                return false;
            }
            if (MonthHeaderWords.Any(w => value.Contains(w)))
            {
                return true;
            }
            return value.Contains('年') && value.Contains('月');
        }

        /// <summary>第 2 列的表头是不是線別类</summary>
        public static bool IsLineHeader(string text)
        {
            string value = (text ?? "").Trim();
            if (value.Length == 0)
            {
                return false;
            }
            return value.Contains("線別") || value.Contains("线别") || value.Contains("線") || value.Contains("线");
        }

        /// <summary>日列表头：「1日」「15日」「01日」</summary>
        public static bool IsDayHeader(string text, out int day)
        {
            day = 0;
            Match m = Regex.Match((text ?? "").Trim(), @"^(\d{1,2})\s*日$");
            if (!m.Success)
            {
                return false;
            }
            if (!int.TryParse(m.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value))
            {
                return false;
            }
            if (value < 1 || value > 31)
            {
                return false;
            }
            day = value;
            return true;
        }

        /// <summary>第 1 列的数据是不是年月（202611 / 2026-11 / 2026/11/1 这类）</summary>
        public static bool IsMonthCell(string text) => PreIssueService.NormalizeMonth(text) != null;

        /// <summary>
        /// 单元格是否像「機種/數量」：有 /，机种部分至少含一个字母、长度够、没有空格。
        /// </summary>
        public static bool LooksLikeModelQty(string text)
        {
            string raw = (text ?? "").Trim();
            int slash = raw.LastIndexOf('/');
            if (slash <= 0 || slash == raw.Length - 1)
            {
                return false;
            }
            string model = raw.Substring(0, slash).Trim();
            string qty = raw.Substring(slash + 1).Trim().Replace(",", "");
            if (model.Length < 3 || !Regex.IsMatch(model, @"^[A-Za-z0-9][A-Za-z0-9\-_.]*$"))
            {
                return false;
            }
            if (!Regex.IsMatch(model, @"[A-Za-z]") || !Regex.IsMatch(model, @"\d"))
            {
                return false;
            }
            return int.TryParse(qty, NumberStyles.Any, CultureInfo.InvariantCulture, out _);
        }

        private static HashSet<string> Normalize(IEnumerable<string> headers)
        {
            HashSet<string> set = new(StringComparer.OrdinalIgnoreCase);
            foreach (string header in headers ?? [])
            {
                string value = (header ?? "").Trim();
                if (value.Length > 0)
                {
                    set.Add(value);
                }
            }
            return set;
        }
    }

    /// <summary>
    /// 导入前的预览：文件是哪一类、涉及哪个月、多少条、会不会覆盖已有数据。
    /// 界面先把它摆给用户看一遍（确认后才真正写库），同时它就是「导错文件」的第一道拦截。
    /// </summary>
    public class PreIssueImportPreview
    {
        /// <summary>来源表（月排程 / ORT看板 / 准备领用）</summary>
        public string Source { get; set; }

        /// <summary>文件名</summary>
        public string File { get; set; }

        /// <summary>能不能导（false 时 <see cref="Message"/> 说明为什么不能）</summary>
        public bool Success { get; set; } = true;

        /// <summary>拦截原因 / 说明</summary>
        public string Message { get; set; }

        /// <summary>涉及月份（yyyyMM）</summary>
        public string Month { get; set; }

        /// <summary>这次会写入的条数</summary>
        public int Items { get; set; }

        /// <summary>涉及多少个机种</summary>
        public int Models { get; set; }

        /// <summary>库里该月已有的条数（导入会覆盖它）</summary>
        public int Existing { get; set; }

        /// <summary>前几条样例（给人核对用）</summary>
        public string Sample { get; set; }

        /// <summary>确认框里显示的文本</summary>
        public string Describe()
        {
            if (!Success)
            {
                return Message ?? "这个文件导不进来";
            }
            string month = PreIssueService.MonthText(Month);
            string sample = string.IsNullOrWhiteSpace(Sample) ? "" : $"\n\n前几条：\n{Sample}";
            string cover = Existing > 0
                ? $"\n（会覆盖 {month} 已有的 {Existing} 条数据；若导错了可以用「撤销上次导入」还原）"
                : "";
            return $"文件：{File}\n月份：{month}\n本次写入：{Items} 条 · {Models} 个机种{cover}{sample}\n\n确认导入吗？";
        }
    }
}

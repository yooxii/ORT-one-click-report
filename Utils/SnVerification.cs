using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ORT一键报告.Utils
{
    /// <summary>
    /// 序列号清单的提取与核对：
    /// - 提取来源：文本（换行/逗号/分号/制表符/顿号分隔）或文件（xls/xlsx/xlsm 自动识别序列号列；csv 按列；txt 按行）；
    /// - 序列号列识别思路：同列非空取值两两不同（序列号不可重复）且存在共同子串（同批序列号共享机型/批次段），
    ///   自动兼容首行表头；多列命中时取（值数最多 → 公共子串最长 → 行列序最前）；
    /// - 核对：报废清单与领用清单必须完全一致（多一个/少一个/清单内重复都算异常）。
    /// </summary>
    public static class SnVerification
    {
        /// <summary>共同子串最少长度（低于该长度不认为是同一批序列号；常量可调）</summary>
        public const int MinCommonSubstring = 3;

        /// <summary>Excel 错误值占位（公式取不到值时导入的脏数据，不参与识别）</summary>
        private static readonly HashSet<string> ErrorPlaceholders = new(StringComparer.OrdinalIgnoreCase)
        {
            "#N/A", "#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!", "#ERROR!"
        };

        /// <summary>文本清单分隔符：换行、半/全角逗号、分号、制表符、顿号</summary>
        private static readonly char[] TextSeparators = ['\r', '\n', ',', '，', ';', '；', '\t', '、'];

        /* ###############################  提取  ################################ */

        /// <summary>
        /// 把用户粘贴的文本拆成序列号列表（去空行与首尾空白，保留顺序与重复，供重复检测）
        /// </summary>
        public static List<string> ParseText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return [];
            }
            return text.Split(TextSeparators, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim())
                .Where(s => s.Length > 0)
                .ToList();
        }

        /// <summary>
        /// 从文件提取序列号清单：按文件头识别 xls/xlsx/xlsm（兼容扩展名写错），
        /// csv 按逗号分列，txt/其他按行取值。识别失败抛异常（信息含原因）。
        /// </summary>
        public static List<string> ExtractFromFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                throw new FileNotFoundException($"文件不存在：{path}");
            }
            byte[] head = ReadHead(path, 8);
            bool zip = head.Length >= 2 && head[0] == 0x50 && head[1] == 0x4B;       // PK → xlsx/xlsm
            bool ole = head.Length >= 4 && head[0] == 0xD0 && head[1] == 0xCF;       // 复合文档 → xls
            if (zip || ole)
            {
                return ExtractFromWorkbook(path);
            }
            string ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext == ".csv")
            {
                return ExtractFromCsv(path);
            }
            // txt 及其它：按行取值（该列仍需满足"两两不同 + 共同子串"）
            return ExtractSingleColumn(ParseText(File.ReadAllText(path)));
        }

        /// <summary>
        /// 领用记录的序列号清单：附件可解析时优先用附件，与 SN 文本并存时取并集
        /// （大小写不敏感去重）；两边都取不到时返回空列表（由调用方提示无法核对）。
        /// </summary>
        public static List<string> ExtractFromRequisition(string snText, string resolvedSnFilePath)
        {
            List<string> fromFile = [];
            if (!string.IsNullOrWhiteSpace(resolvedSnFilePath) && File.Exists(resolvedSnFilePath))
            {
                try
                {
                    fromFile = ExtractFromFile(resolvedSnFilePath);
                }
                catch
                {
                    fromFile = [];  // 附件解析失败时退回 SN 文本（两者都取不到时调用方统一提示）
                }
            }
            List<string> fromText = ParseText(snText);
            if (fromFile.Count == 0)
            {
                return fromText;
            }
            HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);
            List<string> merged = [];
            foreach (string sn in fromFile.Concat(fromText))
            {
                if (seen.Add(sn))
                {
                    merged.Add(sn);
                }
            }
            return merged;
        }

        /* ###############################  核对  ################################ */

        /// <summary>
        /// 核对结果：Ok=true 表示两份清单完全一致
        /// </summary>
        public sealed class SnCompareResult
        {
            /// <summary>是否通过（两份清单非空且完全一致）</summary>
            public bool Ok { get; set; }

            /// <summary>领用清单为空（未登记序列号或附件无法识别）</summary>
            public bool RequisitionListEmpty { get; set; }

            /// <summary>报废清单为空（未提供序列号）</summary>
            public bool ScrapListEmpty { get; set; }

            /// <summary>报废清单有、领用清单没有的序列号</summary>
            public List<string> NotInRequisition { get; set; } = [];

            /// <summary>领用清单有、报废清单没有的序列号（漏报废）</summary>
            public List<string> MissingInScrap { get; set; } = [];

            /// <summary>报废清单内重复的序列号（含出现次数）</summary>
            public List<(string Sn, int Count)> DuplicatesInScrap { get; set; } = [];

            /// <summary>领用清单个数（去重后）</summary>
            public int RequisitionCount { get; set; }

            /// <summary>报废清单个数（去重后）</summary>
            public int ScrapCount { get; set; }
        }

        /// <summary>
        /// 比对领用清单与报废清单：集合完全一致（忽略顺序与大小写）才通过；
        /// "不在领用清单 / 漏报废 / 清单内重复" 任一存在即不通过。
        /// </summary>
        public static SnCompareResult Compare(IReadOnlyList<string> requisition, IReadOnlyList<string> scrap)
        {
            List<string> req = Normalize(requisition);
            List<string> scr = Normalize(scrap);
            List<string> reqDistinct = DistinctIgnoreCase(req);
            List<string> scrDistinct = DistinctIgnoreCase(scr);
            HashSet<string> reqSet = new(reqDistinct, StringComparer.OrdinalIgnoreCase);
            HashSet<string> scrSet = new(scrDistinct, StringComparer.OrdinalIgnoreCase);

            List<(string Sn, int Count)> duplicates = scr
                .GroupBy(s => s, StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() > 1)
                .Select(g => (g.First(), g.Count()))
                .ToList();
            List<string> notInRequisition = scrDistinct.Where(s => !reqSet.Contains(s)).ToList();
            List<string> missingInScrap = reqDistinct.Where(s => !scrSet.Contains(s)).ToList();

            return new SnCompareResult
            {
                RequisitionListEmpty = req.Count == 0,
                ScrapListEmpty = scr.Count == 0,
                NotInRequisition = notInRequisition,
                MissingInScrap = missingInScrap,
                DuplicatesInScrap = duplicates,
                RequisitionCount = reqDistinct.Count,
                ScrapCount = scrDistinct.Count,
                Ok = req.Count > 0 && scr.Count > 0
                    && duplicates.Count == 0 && notInRequisition.Count == 0 && missingInScrap.Count == 0
            };
        }

        private static List<string> Normalize(IReadOnlyList<string> values)
            => values?.Select(v => v?.Trim())
                .Where(v => !string.IsNullOrEmpty(v))
                .ToList() ?? [];

        /// <summary>大小写不敏感去重（保留首次出现的顺序与写法）</summary>
        private static List<string> DistinctIgnoreCase(List<string> values)
        {
            HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);
            List<string> result = [];
            foreach (string value in values)
            {
                if (seen.Add(value))
                {
                    result.Add(value);
                }
            }
            return result;
        }

        /* ###############################  S/N 列识别  ################################ */

        /// <summary>
        /// 从工作簿提取：遍历各表各列，逐列评估候选，取最优
        /// </summary>
        private static List<string> ExtractFromWorkbook(string path)
        {
            IWorkbook workbook = ExcelNpoi.OpenAny(path);
            try
            {
                Candidate best = null;
                for (int s = 0; s < workbook.NumberOfSheets; s++)
                {
                    ISheet sheet = workbook.GetSheetAt(s);
                    int lastRow = ExcelNpoi.LastRow(sheet);
                    int lastCol = ExcelNpoi.LastColumn(sheet);
                    if (lastRow < 1 || lastCol < 1)
                    {
                        continue;
                    }
                    for (int c = 1; c <= lastCol; c++)
                    {
                        List<string> raw = [];
                        for (int r = 1; r <= lastRow; r++)
                        {
                            raw.Add(ExcelNpoi.CellText(sheet, r, c));
                        }
                        Candidate candidate = EvaluateColumn(raw, s, c);
                        if (candidate != null && (best == null || candidate.IsBetterThan(best)))
                        {
                            best = candidate;
                        }
                    }
                }
                if (best == null)
                {
                    throw new InvalidDataException(
                        "未能在文件中识别出序列号列（要求：同一列取值两两不同、且都含共同子串；首行可为表头）");
                }
                return best.Values;
            }
            finally
            {
                workbook.Close();
            }
        }

        /// <summary>
        /// 从 CSV 提取：按行按逗号拆成矩阵后逐列评估（同样只判列，不做行内逗号再拆分）
        /// </summary>
        private static List<string> ExtractFromCsv(string path)
        {
            string[] lines = File.ReadAllLines(path);
            List<string[]> rows = lines.Select(l => l.Split(',')).ToList();
            int colCount = rows.Count == 0 ? 0 : rows.Max(r => r.Length);
            Candidate best = null;
            for (int c = 0; c < colCount; c++)
            {
                List<string> raw = rows.Select(r => c < r.Length ? r[c] : "").ToList();
                Candidate candidate = EvaluateColumn(raw, 0, c);
                if (candidate != null && (best == null || candidate.IsBetterThan(best)))
                {
                    best = candidate;
                }
            }
            if (best == null)
            {
                throw new InvalidDataException(
                    "未能在文件中识别出序列号列（要求：同一列取值两两不同、且都含共同子串；首行可为表头）");
            }
            return best.Values;
        }

        /// <summary>
        /// 单列（txt 按行）提取：仍执行同一套"两两不同 + 共同子串"判定
        /// </summary>
        private static List<string> ExtractSingleColumn(List<string> lines)
        {
            Candidate candidate = EvaluateColumn(lines, 0, 0);
            if (candidate == null)
            {
                throw new InvalidDataException(
                    "未能在文件内容中识别出序列号清单（要求：取值两两不同、且都含共同子串）");
            }
            return candidate.Values;
        }

        /// <summary>
        /// 评估一列原始取值是否是候选序列号列；不合格返回 null。
        /// 规则：单元格内多行先拆开；过滤空白与 Excel 错误占位；非空值 ≥ 2 且两两不同；
        /// 全列存在共同子串（长度 ≥ <see cref="MinCommonSubstring"/>；整列不满足时去掉首行表头重判）。
        /// </summary>
        private static Candidate EvaluateColumn(List<string> rawValues, int sheetIndex, int columnIndex)
        {
            List<string> values = [];
            foreach (string raw in rawValues)
            {
                if (string.IsNullOrWhiteSpace(raw))
                {
                    continue;
                }
                // 单元格内多行（一个单元格放了整份清单的情况）逐个拆出
                foreach (string part in raw.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries))
                {
                    string trimmed = part.Trim();
                    if (trimmed.Length > 0 && !ErrorPlaceholders.Contains(trimmed))
                    {
                        values.Add(trimmed);
                    }
                }
            }
            if (values.Count < 2)
            {
                return null;
            }
            if (DistinctIgnoreCase(values).Count != values.Count)
            {
                return null; // 存在重复：不是序列号列（序列号不可重复）
            }
            int common = LongestCommonSubstringLength(values);
            if (common < MinCommonSubstring && values.Count >= 3)
            {
                // 首行可能是表头（如 "S/N"）：去掉首行后重判
                List<string> stripped = values.Skip(1).ToList();
                if (DistinctIgnoreCase(stripped).Count == stripped.Count)
                {
                    int strippedCommon = LongestCommonSubstringLength(stripped);
                    if (strippedCommon >= MinCommonSubstring)
                    {
                        values = stripped;
                        common = strippedCommon;
                    }
                }
            }
            if (common < MinCommonSubstring)
            {
                return null;
            }
            return new Candidate { Values = values, CommonLength = common, SheetIndex = sheetIndex, ColumnIndex = columnIndex };
        }

        /// <summary>
        /// 候选列：值数最多 → 公共子串最长 → 工作表/列序最前
        /// </summary>
        private sealed class Candidate
        {
            public List<string> Values;
            public int CommonLength;
            public int SheetIndex;
            public int ColumnIndex;

            public bool IsBetterThan(Candidate other)
            {
                if (Values.Count != other.Values.Count)
                {
                    return Values.Count > other.Values.Count;
                }
                if (CommonLength != other.CommonLength)
                {
                    return CommonLength > other.CommonLength;
                }
                return SheetIndex != other.SheetIndex ? SheetIndex < other.SheetIndex : ColumnIndex < other.ColumnIndex;
            }
        }

        /// <summary>
        /// 全部取值的公共子串最大长度（不足 <see cref="MinCommonSubstring"/> 时返回 0）。
        /// 以最短的字符串为基准枚举子串（任意公共子串必然是其子串），逐长检查是否被所有取值包含。
        /// </summary>
        public static int LongestCommonSubstringLength(IReadOnlyList<string> values)
        {
            if (values == null || values.Count == 0)
            {
                return 0;
            }
            string basis = values.OrderBy(v => v.Length).First();
            for (int len = basis.Length; len >= MinCommonSubstring; len--)
            {
                for (int start = 0; start + len <= basis.Length; start++)
                {
                    string piece = basis.Substring(start, len);
                    bool all = true;
                    foreach (string value in values)
                    {
                        if (value.IndexOf(piece, StringComparison.OrdinalIgnoreCase) < 0)
                        {
                            all = false;
                            break;
                        }
                    }
                    if (all)
                    {
                        return len;
                    }
                }
            }
            return 0;
        }

        private static byte[] ReadHead(string path, int count)
        {
            try
            {
                using FileStream fs = File.OpenRead(path);
                byte[] buffer = new byte[count];
                int read = fs.Read(buffer, 0, count);
                return read == count ? buffer : buffer.Take(read).ToArray();
            }
            catch
            {
                return [];
            }
        }
    }
}

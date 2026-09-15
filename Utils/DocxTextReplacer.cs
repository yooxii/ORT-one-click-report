using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ORT一键报告.Utils
{
    /// <summary>Word 文本替换结果</summary>
    public class DocxReplaceResult
    {
        public int Files { get; set; }
        public int ChangedFiles { get; set; }
        public int Replacements { get; set; }
        public int Failed { get; set; }
        public List<string> Messages { get; } = [];
    }

    /// <summary>
    /// Word 文档字符串批量替换：直接用 OpenXML SDK 改文档 XML 再保存。
    /// 关键点：只重写文本节点，**不会重新编码/压缩图片**（图片部件原样复制），
    /// 因此比"用 Word 打开另存"更快也更安全，适合一次替换很多组字符串。
    /// </summary>
    public static class DocxTextReplacer
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// 解析"旧=新"逐行文本为替换对（忽略空行与 # 开头的注释行；分隔符支持 = 和 →）
        /// </summary>
        public static List<KeyValuePair<string, string>> ParsePairs(string text)
        {
            List<KeyValuePair<string, string>> pairs = [];
            foreach (string raw in (text ?? "").Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries))
            {
                string line = raw.Trim();
                if (line.Length == 0 || line.StartsWith("#"))
                {
                    continue;
                }
                int idx = line.IndexOf('=');
                if (idx < 0)
                {
                    idx = line.IndexOf('→');
                }
                if (idx <= 0)
                {
                    continue;
                }
                string oldValue = line.Substring(0, idx).Trim();
                string newValue = line.Substring(idx + 1).Trim();
                if (oldValue.Length == 0)
                {
                    continue;
                }
                pairs.Add(new KeyValuePair<string, string>(oldValue, newValue));
            }
            return pairs;
        }

        /// <summary>
        /// 批量替换目录下所有 .docx（含子目录），默认先备份为 *.docx.bak（已存在则不覆盖）
        /// </summary>
        public static DocxReplaceResult ReplaceInFolder(string folder, IEnumerable<KeyValuePair<string, string>> pairs, bool backup = true)
        {
            DocxReplaceResult result = new();
            List<KeyValuePair<string, string>> list = pairs?.Where(p => !string.IsNullOrEmpty(p.Key)).ToList() ?? [];
            if (list.Count == 0)
            {
                result.Messages.Add("没有可用的替换规则（请按 旧=新 每行一组填写）");
                return result;
            }
            if (!Directory.Exists(folder))
            {
                result.Messages.Add($"目录不存在：{folder}");
                return result;
            }
            foreach (string file in Directory.GetFiles(folder, "*.docx", SearchOption.AllDirectories))
            {
                result.Files++;
                try
                {
                    int count = ReplaceInFile(file, list, backup);
                    if (count > 0)
                    {
                        result.ChangedFiles++;
                        result.Replacements += count;
                    }
                }
                catch (Exception ex)
                {
                    result.Failed++;
                    result.Messages.Add($"{Path.GetFileName(file)}：{ex.Message}");
                    _logger.Error(ex, $"Word 文本替换失败：{file}");
                }
            }
            _logger.Info($"Word 文本替换完成：文件 {result.Files} 个，改动 {result.ChangedFiles} 个，替换 {result.Replacements} 处，失败 {result.Failed} 个（{folder}）");
            return result;
        }

        /// <summary>
        /// 替换单个 Word 文档中的字符串，返回替换处数
        /// </summary>
        public static int ReplaceInFile(string path, IEnumerable<KeyValuePair<string, string>> pairs, bool backup = true)
        {
            List<KeyValuePair<string, string>> list = pairs?.Where(p => !string.IsNullOrEmpty(p.Key)).ToList() ?? [];
            if (list.Count == 0)
            {
                return 0;
            }
            if (backup)
            {
                string backupPath = path + ".bak";
                if (!File.Exists(backupPath))
                {
                    File.Copy(path, backupPath, false);
                }
            }

            int replacements = 0;
            using (WordprocessingDocument doc = WordprocessingDocument.Open(path, true))
            {
                replacements += ReplaceInRoot(doc.MainDocumentPart?.Document?.Body, list);
                if (doc.MainDocumentPart != null)
                {
                    foreach (HeaderPart header in doc.MainDocumentPart.HeaderParts)
                    {
                        replacements += ReplaceInRoot(header.Header, list);
                    }
                    foreach (FooterPart footer in doc.MainDocumentPart.FooterParts)
                    {
                        replacements += ReplaceInRoot(footer.Footer, list);
                    }
                }
                doc.MainDocumentPart?.Document?.Save();
            }
            return replacements;
        }

        private static int ReplaceInRoot(DocumentFormat.OpenXml.OpenXmlElement root, List<KeyValuePair<string, string>> pairs)
        {
            if (root == null)
            {
                return 0;
            }
            int count = 0;
            foreach (Paragraph paragraph in root.Descendants<Paragraph>())
            {
                count += ReplaceInParagraph(paragraph, pairs);
            }
            return count;
        }

        /// <summary>
        /// 段落内替换：
        /// 1) 先逐 run 直接替换（覆盖绝大多数情况，且完全保留原有格式）；
        /// 2) 若同一段落里还有跨 run 断开的字符串，再做一次"拼起来替换后写回首個 run"的处理（只影响该段）。
        /// </summary>
        private static int ReplaceInParagraph(Paragraph paragraph, List<KeyValuePair<string, string>> pairs)
        {
            int count = 0;
            List<Text> texts = paragraph.Descendants<Text>().ToList();
            if (texts.Count == 0)
            {
                return 0;
            }

            // 1) run 内替换
            foreach (Text text in texts)
            {
                string value = text.Text;
                if (string.IsNullOrEmpty(value))
                {
                    continue;
                }
                string replaced = value;
                foreach (KeyValuePair<string, string> pair in pairs)
                {
                    if (replaced.IndexOf(pair.Key, StringComparison.Ordinal) >= 0)
                    {
                        count += CountOccurrences(replaced, pair.Key);
                        replaced = replaced.Replace(pair.Key, pair.Value);
                    }
                }
                if (!ReferenceEquals(replaced, value) && replaced != value)
                {
                    text.Text = replaced;
                    text.Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve;
                }
            }

            // 2) 跨 run 替换
            string joined = string.Concat(texts.Select(t => t.Text));
            string joinedNew = joined;
            foreach (KeyValuePair<string, string> pair in pairs)
            {
                if (joinedNew.IndexOf(pair.Key, StringComparison.Ordinal) >= 0)
                {
                    joinedNew = joinedNew.Replace(pair.Key, pair.Value);
                }
            }
            if (joinedNew != joined)
            {
                int crossCount = 0;
                string probe = joined;
                foreach (KeyValuePair<string, string> pair in pairs)
                {
                    int occurrences = CountOccurrences(probe, pair.Key);
                    if (occurrences > 0)
                    {
                        crossCount += occurrences;
                        probe = probe.Replace(pair.Key, pair.Value);
                    }
                }
                texts[0].Text = joinedNew;
                texts[0].Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve;
                for (int i = 1; i < texts.Count; i++)
                {
                    texts[i].Text = "";
                }
                count += crossCount;
            }
            return count;
        }

        private static int CountOccurrences(string text, string needle)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(needle))
            {
                return 0;
            }
            int count = 0, index = 0;
            while ((index = text.IndexOf(needle, index, StringComparison.Ordinal)) >= 0)
            {
                count++;
                index += needle.Length;
            }
            return count;
        }
    }
}

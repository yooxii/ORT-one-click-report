using NLog;
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ORT一键报告.Utils
{
    /// <summary>
    /// 生成后的 xlsx 收尾清理。
    ///
    /// 报告模板是从历史报告另存来的，里面残留了大量"无效记录"：外部工作簿引用
    /// （xl/externalLinks/*，缓存值早已失效）、失效的命名区域（600+ 个 #REF! 名字）、
    /// 以及指向已删除工作表的 calcChain。NPOI 会把这些原样写回，Excel 打开时就报
    /// "检测到错误…已删除记录：命名区域(工作簿) / 外部公式引用 / 共享公式 / 公式(计算属性)"
    /// 并要求修复。
    ///
    /// 这里在保存之后直接改 xlsx 压缩包：删掉这几类部件、声明与关系，
    /// 让生成的文件成为干净的工作簿（工作表内容、图片、超链接都不受影响）。
    /// </summary>
    public static class XlsxCleaner
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private static readonly Regex DefinedNameRegex = new(@"<definedName\b[^>]*>.*?</definedName>", RegexOptions.Singleline);
        private static readonly Regex ExternalReferencesRegex = new(@"<externalReferences\b[^>]*>.*?</externalReferences>", RegexOptions.Singleline);
        private static readonly Regex RelationshipRegex = new(@"<Relationship\b[^>]*/>");
        private static readonly Regex OverrideRegex = new(@"<Override\b[^>]*/>");
        private static readonly Regex CalcPrRegex = new(@"<calcPr\b[^>]*/>|<calcPr\b[^>]*>.*?</calcPr>", RegexOptions.Singleline);

        /// <summary>
        /// 清理生成出来的报告概览文件；失败只记日志，不影响生成结果
        /// </summary>
        public static bool Clean(string xlsxPath)
        {
            if (string.IsNullOrWhiteSpace(xlsxPath) || !File.Exists(xlsxPath))
            {
                return false;
            }
            try
            {
                using FileStream stream = new(xlsxPath, FileMode.Open, FileAccess.ReadWrite);
                using ZipArchive zip = new(stream, ZipArchiveMode.Update);

                // 1. 删除失效部件：calcChain（行删/插后条目与工作表对不上）、外部工作簿引用（缓存值已失效）
                foreach (ZipArchiveEntry entry in zip.Entries.Where(IsRemovablePart).ToList())
                {
                    entry.Delete();
                }

                // 2. workbook.xml：去掉失效的命名区域与外部引用声明，并要求打开时全量重算
                Rewrite(zip, "xl/workbook.xml", text => ForceFullCalc(
                    ExternalReferencesRegex.Replace(DefinedNameRegex.Replace(text, m => IsInvalidDefinedName(m.Value) ? "" : m.Value), "")));

                // 3. 关系：去掉 calcChain / externalLink 的关系项
                Rewrite(zip, "xl/_rels/workbook.xml.rels", text => RelationshipRegex.Replace(text, m =>
                    m.Value.Contains("/calcChain") || m.Value.Contains("/externalLink") ? "" : m.Value));

                // 4. 内容类型声明：去掉已删除部件的 Override
                Rewrite(zip, "[Content_Types].xml", text => OverrideRegex.Replace(text, m =>
                    m.Value.Contains("/xl/calcChain.xml") || m.Value.Contains("/xl/externalLinks/") ? "" : m.Value));

                return true;
            }
            catch (Exception ex)
            {
                Logger.Warn($"清理生成文件的无效记录失败（{xlsxPath}）：{ex.Message}");
                return false;
            }
        }

        /// <summary>是否是要删掉的部件（calcChain 或外部工作簿引用）</summary>
        private static bool IsRemovablePart(ZipArchiveEntry entry)
        {
            string name = entry.FullName;
            return name.EndsWith("calcChain.xml", StringComparison.OrdinalIgnoreCase)
                || name.StartsWith("xl/externalLinks/", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>命名区域是否失效（引用 #REF! 或指向外部工作簿）</summary>
        private static bool IsInvalidDefinedName(string element)
        {
            return element.Contains("#REF!") || element.Contains("[");
        }

        /// <summary>
        /// 让 Excel 打开时全量重算：写出的公式没有缓存值，不重算的话可能显示成 0
        /// </summary>
        private static string ForceFullCalc(string workbookXml)
        {
            const string calcPr = "<calcPr fullCalcOnLoad=\"1\"/>";
            if (CalcPrRegex.IsMatch(workbookXml))
            {
                return CalcPrRegex.Replace(workbookXml, calcPr, 1);
            }
            int index = workbookXml.IndexOf("</workbook>", StringComparison.Ordinal);
            return index < 0 ? workbookXml : workbookXml.Insert(index, calcPr);
        }

        /// <summary>
        /// 重写压缩包里的一个文本部件（先读出内容，再删掉重建，避免流的长度限制）
        /// </summary>
        private static void Rewrite(ZipArchive zip, string entryName, Func<string, string> transform)
        {
            ZipArchiveEntry entry = zip.GetEntry(entryName);
            if (entry == null)
            {
                return;
            }
            string text;
            using (StreamReader reader = new(entry.Open(), new UTF8Encoding(false)))
            {
                text = reader.ReadToEnd();
            }
            string updated = transform(text);
            if (string.Equals(text, updated, StringComparison.Ordinal))
            {
                return;
            }
            entry.Delete();
            ZipArchiveEntry created = zip.CreateEntry(entryName, CompressionLevel.Optimal);
            using StreamWriter writer = new(created.Open(), new UTF8Encoding(false));
            writer.Write(updated);
        }
    }
}

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ORT一键报告.Utils
{
    public class SmartTableExtractor
    {
        // 定义一个通用的表格数据结构
        // 注意：为了容纳"嵌套表格"，我们将单元格的类型从 string 改为 object
        public class CellContent
        {
            public string Text { get; set; } // 普通文本
            public List<List<CellContent>> NestedTable { get; set; } // 嵌套的表格（如果存在）

            // 用于判断是否为空
            public bool IsEmpty => string.IsNullOrWhiteSpace(Text) && NestedTable == null;
        }

        /// <summary>
        /// 将 Word 文档中所有表格（包含嵌套表格）提取并转换为 CSV 字符串
        /// </summary>
        /// <param name="filePath">Word文档路径</param>
        /// <returns>完整的 CSV 字符串</returns>
        public static string ConvertWordTablesToCsv(string filePath)
        {
            var csvOutput = new StringBuilder();

            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
            {
                var body = doc.MainDocumentPart.Document.Body;
                var tables = body.Elements<Table>();

                foreach (var table in tables)
                {
                    // 提取并转换当前表格
                    var tableData = ExtractTableRecursive(table);
                    ConvertTableDataToCsvString(tableData, csvOutput);

                    // 每个表格之间加一个空行，方便区分
                    csvOutput.AppendLine();
                }
            }

            return csvOutput.ToString();
        }

        /// <summary>
        /// 递归提取表格（能处理嵌套）
        /// </summary>
        public static List<List<CellContent>> ExtractTableRecursive(Table table)
        {
            var tableData = new List<List<CellContent>>();

            foreach (var row in table.Elements<TableRow>())
            {
                var rowData = new List<CellContent>();

                foreach (var cell in row.Elements<TableCell>())
                {
                    var cellContent = new CellContent();

                    // 1. 检查单元格内是否包含嵌套表格
                    var nestedTable = cell.Elements<Table>().FirstOrDefault();

                    if (nestedTable != null)
                    {
                        // 2. 如果包含表格，递归提取该表格
                        cellContent.NestedTable = ExtractTableRecursive(nestedTable);
                        // 注意：此时我们通常忽略该单元格内的纯文本（或者可以同时保留文本和表格，视需求而定）
                        // 这里为了结构清晰，优先提取表格，忽略同级文本
                    }
                    else
                    {
                        // 3. 如果不包含表格，提取纯文本
                        cellContent.Text = ExtractTextFromCell(cell);
                    }

                    rowData.Add(cellContent);
                }
                tableData.Add(rowData);
            }

            return tableData;
        }

        /// <summary>
        /// 将提取出的表格数据递归转换为 CSV 格式字符串
        /// </summary>
        private static void ConvertTableDataToCsvString(List<List<CellContent>> tableData, StringBuilder sb)
        {
            foreach (var row in tableData)
            {
                var csvRow = new List<string>();
                foreach (var cell in row)
                {
                    if (cell.NestedTable != null)
                    {
                        // 如果单元格是嵌套表格，将其递归转换为字符串后，整体作为一个单元格内容
                        var nestedSb = new StringBuilder();
                        ConvertTableDataToCsvString(cell.NestedTable, nestedSb);
                        // 去掉末尾多余的换行符，防止破坏外层 CSV 结构
                        var nestedCsv = nestedSb.ToString().TrimEnd('\r', '\n');
                        csvRow.Add(EscapeCsvField(nestedCsv));
                    }
                    else
                    {
                        csvRow.Add(EscapeCsvField(cell.Text ?? ""));
                    }
                }
                // 将当前行的所有单元格用逗号拼接，并追加到总输出中
                sb.AppendLine(string.Join(",", csvRow));
            }
        }

        /// <summary>
        /// CSV 字段转义核心方法（遵循 RFC 4180 规范）
        /// </summary>
        private static string EscapeCsvField(string field)
        {
            // 如果字段包含 逗号(,)、双引号(") 或 换行符(\n, \r)，必须用双引号包裹
            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
            {
                // 字段内部的双引号，需要转义为两个连续的双引号 ("")
                return "\"" + field.Replace("\"", "\"\"") + "\"";
            }
            return field;
        }

        /// <summary>
        /// 提取单元格纯文本（兼容之前的逻辑，用于非嵌套场景）
        /// </summary>
        private static string ExtractTextFromCell(TableCell cell)
        {
            var texts = cell.Descendants<Text>().Select(t => t.Text);
            return string.Join("", texts);
        }
    }
}
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NLog;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace ORT一键报告
{
    public class SmartTableExtractor
    {
        // 定义一个通用的表格数据结构
        // 注意：为了容纳“嵌套表格”，我们将单元格的类型从 string 改为 object
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
            var texts = cell.Descendants<Text>()
                           .Select(t => t.Text);
            return string.Join("", texts);
        }
    }

    public class ReportUtils
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public class DataCell
        {
            public string Data { get; set; }
            public int Row { get; set; }
            public int Column { get; set; }
            public List<ExcelPictureInfo> Images { get; set; } = null;
            public string TopLeftAddress
            {
                get => ExcelCellBase.GetAddress(Row, Column);
                set
                {
                    int bRow = Row;
                    int bColumn = Column;
                    try
                    {
                        ExcelAddress Addr = new ExcelAddress(value);
                        Row = Addr.Start.Row;
                        Column = Addr.Start.Column;
                    }
                    catch
                    {
                        Row = bRow;
                        Column = bColumn;
                    }
                }
            }
            public override string ToString()
            {
                return $"{Data} - {TopLeftAddress}({Row},{Column})";
            }
        }

        public class TestItemInfo
        {
            public string TestItemName { get; set; }
            public string Date { get; set; }
        }

        public class UUTInfoFromExcel
        {
            public List<string> SNs { get; set; }
            public string WorkerNo { get; set; }
            public string Revision { get; set; }
            public string DC { get; set; }
            public List<TestItemInfo> TestItems { get; set; }

            public override string ToString()
            {
                return $"{WorkerNo},{Revision},{DC},{(TestItems == null ? 0 : TestItems.Count)},{(SNs == null ? 0 : SNs.Count)}";
            }
        }

        public class ReportHeaderInfo
        {
            public DataCell TESTED_BY { get; set; }
            public DataCell APPROVED_BY { get; set; }
            public DataCell PROJECT_NAME { get; set; }
            public DataCell TEST_STAGE { get; set; }
            public DataCell TestDescription { get; set; }
            public DataCell Test_Description_Pic { get; set; }
            public DataCell Issue_Photos_Pics { get; set; }
            public DataCell Test_Setup_Pics { get; set; }
            public DataCell Test_ATE_Data { get; set; }
        }

        public class ResultDetails
        {
            public string BIroom { get; set; } = "";
            public string BIarea { get; set; } = "";
            public string BIplace { get; set; } = "";
            public string SN { get; set; } = "";
            public string WorkOrder { get; set; } = "";
            public string Version { get; set; } = "";
            public string DC { get; set; } = "";
            public ReportStatus InspectionPrev { get; set; }
            public ReportStatus FunPrev { get; set; }
            public ReportStatus InspectionAfter { get; set; }
            public ReportStatus FunAfter { get; set; }
            public ReportStatus HiPot { get; set; }
            public string Comments { get; set; } = "";
        }

        /// <summary>
        /// 辅助类：用于返回提取的图片信息
        /// </summary>
        public class ExcelPictureInfo
        {
            public ExcelPicture Picture { get; set; } // 原始对象
            public ImageSource ImageSrc { get; set; }    // System.Drawing.Image 对象
            public byte[] ImageBytes { get; set; }    // 字节数组
            public string Name { get; set; }          // 图片名称
        }


        /* ###############################  EPPlus函数  ################################ */

        public static DataCell FindCellByValue(ExcelWorksheet ws, string value, string excludeValue = "")
        {
            int snRowStart = 1;
            int snColumnStart = 1;
            int snColumnEnd = ws.Dimension.End.Column;
            int snRowEnd = ws.Dimension.End.Row;
            value = value.ToLower();
            excludeValue = excludeValue.ToLower();

            DataCell result;
            for (int row = snRowStart; row <= snRowEnd; row++)
            {
                for (int col = snColumnStart; col <= snColumnEnd; col++)
                {
                    var _value = ws.Cells[row, col].Text;
                    if (_value.ToLower().Contains(value))
                    {
                        if (_value.ToLower().Contains(excludeValue) && excludeValue != "")
                        {
                            continue;
                        }
                        result = new DataCell
                        {
                            Data = _value,
                            Row = row,
                            Column = col
                        };
                        return result;
                    }
                }
            }
            return null;
        }

        public static void EmbedOleObjectWithInterop(string targetExcelPath, string objectToEmbedPath, string TopLeftAddress = "A1")
        {
            _logger.Info($"插入OLE对象到{targetExcelPath}...");
            if (objectToEmbedPath is null || objectToEmbedPath == "")
            {
                _logger.Warn($"OLE对象路径({objectToEmbedPath})为空");
                return;
            }
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Microsoft.Office.Interop.Excel.Workbook workbook = null;
            Microsoft.Office.Interop.Excel.Worksheet worksheet = null;

            try
            {
                // 1. 启动 Excel 应用
                excelApp = new Microsoft.Office.Interop.Excel.Application
                {
                    Visible = true,
                    DisplayAlerts = false
                };

                // 2. 打开目标文件
                workbook = excelApp.Workbooks.Open(targetExcelPath);
                worksheet = workbook.Worksheets[1];

                // 3. 定义嵌入位置 (例如 A1 单元格)
                Microsoft.Office.Interop.Excel.Range range = worksheet.Range[TopLeftAddress];
                double left = range.Left;
                double top = range.Top;

                // 4. 执行嵌入操作
                worksheet.OLEObjects().Add(
                    Filename: objectToEmbedPath,
                    Link: false,
                    DisplayAsIcon: true,
                    IconFileName: Type.Missing,
                    IconIndex: Type.Missing,
                    IconLabel: "点击查看详细数据",
                    Left: left,
                    Top: top
                );

                // 5. 保存并关闭
                workbook.Save();
                workbook.Close();
                _logger.Info("OLE对象插入成功");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "OLE对象插入失败");
            }
            finally
            {
                // 6. 清理 COM 对象 (非常重要，防止内存泄漏)
                if (workbook != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public static ReportHeaderInfo ReadReportHeaderInfo(ExcelWorksheet ws)
        {
            // 辅助函数: 找到issue和setup图片所在的标题行
            DataCell issueTitle = FindCellByValue(ws, "Issue Photos");
            DataCell setupTitle = FindCellByValue(ws, "Test Setup");

            ReportHeaderInfo reportHeaderInfo = new ReportHeaderInfo
            {
                TESTED_BY = FindInfoByText(ws, "TESTED BY"),
                APPROVED_BY = FindInfoByText(ws, "APPROVED BY"),
                PROJECT_NAME = FindInfoByText(ws, "PROJECT NAME"),
                TEST_STAGE = FindInfoByText(ws, "TEST STAGE"),
                TestDescription = FindInfoByText(ws, "Test Description"),
                Test_Description_Pic = GetPicturesInRange(ws, 6, 1, 10),
                Issue_Photos_Pics = issueTitle is null ? null : GetPicturesInRange(ws, issueTitle.Row, 1, issueTitle.Row + 10),
                Test_Setup_Pics = setupTitle is null ? null : GetPicturesInRange(ws, setupTitle.Row, 1, setupTitle.Row + 10),
            };
            return reportHeaderInfo;
        }

        public static DataCell GetPicturesInRange(ExcelWorksheet ws, int startRow = 1, int startCol = 1, int endRow = -1, int endCol = -1)
        {
            if (ws == null || ws.Drawings.Count == 0)
            {
                return null;
            }

            var result = new DataCell()
            {
                Images = new List<ExcelPictureInfo>()
            };

            if (endRow == -1)
            {
                endRow = ws.Dimension.End.Row;
            }
            if (endCol == -1)
            {
                endCol = ws.Dimension.End.Column;
            }

            // 规范化范围 (防止用户传反了行列)
            int minRow = Math.Min(startRow, endRow);
            int maxRow = Math.Max(startRow, endRow);
            int minCol = Math.Min(startCol, endCol);
            int maxCol = Math.Max(startCol, endCol);

            foreach (var drawing in ws.Drawings)
            {
                if (drawing is ExcelPicture picture)
                {
                    // 获取图片左上角锚定的单元格坐标
                    int picRow = picture.From.Row + 1; // EPPlus Row 索引从 0 开始，Excel 从 1 开始
                    int picCol = picture.From.Column + 1;

                    // 判断逻辑：只要图片的左上角在指定范围内，就视为在该范围内
                    if (picRow >= minRow && picRow <= maxRow &&
                        picCol >= minCol && picCol <= maxCol)
                    {
                        result.Images.Add(new ExcelPictureInfo()
                        {
                            Picture = picture,
                            ImageSrc = ConvertToWpfImage(picture.Image.ImageBytes),
                            ImageBytes = picture.Image.ImageBytes,
                            Name = picture.Name,
                        });
                        result.Data = "Images";
                        result.Row = picRow;
                        result.Column = picCol;
                    }
                }
            }
            result.Images.Reverse();
            return result;
        }

        public static DataCell FindInfoByText(ExcelWorksheet ws, string toFind)
        {
            DataCell headerInfo = new DataCell();
            var cell = FindCellByValue(ws, toFind);
            if (cell != null)
            {
                for (int c = cell.Column + 1; c <= ws.Dimension.End.Column; c++)
                {
                    string value = ws.Cells[cell.Row, c].Text;
                    if (value != "")
                    {
                        headerInfo.Data = value;
                        headerInfo.Row = cell.Row;
                        headerInfo.Column = c;
                        break;
                    }
                }
            }
            return headerInfo;
        }


        /* ###############################  功能函数  ################################ */

        public static string GetTemplatePath(string rootPath, string reportType)
        {
            string[] excelExtensions = new[] { ".xlsx", ".xls", ".xlsm" };
            string[] excelFiles = Directory.GetFiles(rootPath, "*.*", SearchOption.AllDirectories).Where(file => excelExtensions.Contains(Path.GetExtension(file))).ToArray();
            Regex regex = new Regex(@"[^a-zA-Z]");
            foreach (string excelFile in excelFiles)
            {
                if (regex.Replace(excelFile, "").ToLower().Contains(regex.Replace(reportType, "").ToLower()))
                {
                    return excelFile;
                }
            }
            return "";
        }

        public static string GetSubstringAfter(string source, string marker, int length)
        {
            if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(marker))
            {
                return string.Empty;
            }

            int index = source.IndexOf(marker);
            if (index == -1) // 未找到标记
            {
                return string.Empty;
            }

            int startIndex = index + marker.Length;
            if (startIndex >= source.Length)
            {
                return string.Empty;
            }

            int actualLength = Math.Min(length, source.Length - startIndex);
            return source.Substring(startIndex, actualLength);
        }

        /// <summary>
        /// 将图片字节数组转换为 BitmapImage
        /// </summary>
        public static ImageSource ConvertToWpfImage(byte[] imageBytes)
        {
            if (imageBytes == null || imageBytes.Length == 0)
            {
                return null;
            }

            var bitmapImage = new BitmapImage();
            using (var ms = new MemoryStream(imageBytes))
            {
                bitmapImage.BeginInit();
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad; // 重要：加载后释放流
                bitmapImage.StreamSource = ms;
                bitmapImage.EndInit();
                bitmapImage.Freeze(); // 冻结以提高性能并允许跨线程访问
            }
            return bitmapImage;
        }

        public static void ClearTempDir()
        {
            _logger.Info("清理临时目录...");
            string TempPath = Path.Combine(Path.GetTempPath(), "ORTTemp");
            try
            {
                foreach (string fl in Directory.GetFiles(TempPath))
                {
                    File.Delete(fl);
                }
                Directory.Delete(TempPath);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "清理失败");
            }
            _logger.Info("清理完成");
        }
    }
}

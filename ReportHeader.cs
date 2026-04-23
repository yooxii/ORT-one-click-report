using NLog;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace ORT一键报告
{
    public class ReportHeader
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

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


        public static void EmbedOleObjectWithInterop(Logger logger, string targetExcelPath, string objectToEmbedPath, string TopLeftAddress = "A1")
        {
            logger.Info("插入OLE对象...");
            if (objectToEmbedPath is null || objectToEmbedPath == "")
            {
                logger.Warn("OLE对象路径为空");
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
                logger.Info("OLE对象插入成功");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "OLE对象插入失败");
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

        public static string GetSubstringAfter(string source, string marker, int length)
        {
            if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(marker))
                return string.Empty;

            int index = source.IndexOf(marker);
            if (index == -1) // 未找到标记
                return string.Empty;

            int startIndex = index + marker.Length;
            if (startIndex >= source.Length)
                return string.Empty;

            int actualLength = Math.Min(length, source.Length - startIndex);
            return source.Substring(startIndex, actualLength);
        }

        /// <summary>
        /// 将 EPPlus 的图片字节数组转换为 WPF 可用的 BitmapImage
        /// </summary>
        public static ImageSource ConvertToWpfImage(byte[] imageBytes)
        {
            if (imageBytes == null || imageBytes.Length == 0)
                return null;

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

        public static string GetTemplatePath(string rootPath, string reportTag)
        {
            string[] excelExtensions = new[] { ".xlsx", ".xls", ".xlsm" };
            string[] excelFiles = Directory.GetFiles(rootPath, "*.*", SearchOption.AllDirectories).Where(file => excelExtensions.Contains(System.IO.Path.GetExtension(file))).ToArray();
            foreach (string excelFile in excelFiles)
            {
                if (excelFile.ToLower().Contains(reportTag.ToLower()))
                {
                    return excelFile;
                }
            }
            return "";
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

        public static void ClearTempDir(Logger logger)
        {
            logger.Info("清理临时目录...");
            string TempPath = Path.Combine(Path.GetTempPath(), "ORTTemp");
            try
            {
                foreach (var fl in Directory.GetFiles(TempPath))
                {
                    File.Delete(fl);
                }
                Directory.Delete(TempPath);
            }
            catch (Exception ex)
            {
                logger.Error(ex, "清理失败");
            }
            logger.Info("清理完成");
        }
    }
}

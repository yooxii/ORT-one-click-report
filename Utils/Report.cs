using NLog;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.OleObject;
using ORT一键报告.Models;
using ORT一键报告.Reports.ViewModels;
using ORT一键报告.Reports.Views;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;

namespace ORT一键报告.Utils
{
    public class Report
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();


        #region EPPlus

        public static string GetCellAddress(int row, int column)
        {
            return ExcelCellBase.GetAddress(row, column);
        }
        public static string GetCellColumn(int column)
        {
            return ExcelCellBase.GetAddress(1, column).Replace("1", "");
        }

        public static DataCell FindCellByValue(ExcelWorksheet ws, string value, string excludeValue = "", bool ignoreCase = true, DataCell start = null, DataCell end = null)
        {
            int snRowStart = 1;
            int snColumnStart = 1;
            int snColumnEnd = ws.Dimension.End.Column;
            int snRowEnd = ws.Dimension.End.Row;
            DataCell result;

            if (start != null)
            {
                snRowStart = start.Row;
                snColumnStart = start.Column;
            }
            if (end != null)
            {
                snRowEnd = end.Row;
                snColumnEnd = end.Column;
            }

            if (snRowEnd < snRowStart || snColumnEnd < snColumnStart)
            {
                _logger.Warn("搜索的范围过小！");
                return null;
            }

            if (ignoreCase)
            {
                value = value.ToLower();
                excludeValue = excludeValue.ToLower();
            }

            for (int row = snRowStart; row <= snRowEnd; row++)
            {
                for (int col = snColumnStart; col <= snColumnEnd; col++)
                {
                    var _value = ws.Cells[row, col].Text;
                    if (ignoreCase)
                        _value = _value.ToLower();
                    if (_value.Contains(value))
                    {
                        if (excludeValue != "" && _value.Contains(excludeValue))
                        {
                            continue;
                        }
                        result = new DataCell(row, col) { Data = _value };
                        return result;
                    }
                }
            }
            return null;
        }

        public static void EmbedOleObjectWithInterop(string targetExcelPath, string objectToEmbedPath, string TopLeftAddress = "A1")
        {
            _logger.Info($"插入OLE对象到{targetExcelPath}...");
            if (objectToEmbedPath is null or "")
            {
                _logger.Warn($"OLE对象路径({objectToEmbedPath})为空");
                return;
            }
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Microsoft.Office.Interop.Excel.Workbook workbook = null;
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
                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];

                // 3. 定义嵌入位置 (例如 A1 单元格)
                Microsoft.Office.Interop.Excel.Range range = worksheet.Range[TopLeftAddress];
                double left = (double)range.Left;
                double top = (double)range.Top;

                // 4. 执行嵌入操作
                dynamic oleObjects = worksheet.OLEObjects(); // 提前获取 OLE 对象集合
                oleObjects.Add(
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

        public static void EmbedOleObjectWithEpplus(ExcelWorksheet ws, string objectToEmbedPath, string TopLeftAddress = "A1", string IconPath = "", int IconX = 10, int IcnoY = 10, int IconW = 100, int IconH = 100)
        {
            _logger.Info($"插入OLE对象到{ws.Name}...");
            if (objectToEmbedPath is null or "")
            {
                _logger.Warn($"OLE对象路径({objectToEmbedPath})为空");
                return;
            }

            using MemoryStream iconStream = new(Resources.image_xlsx_emf);
            iconStream.Position = 0; // 必须重置流指针到开头
            try
            {
                DataCell tmp = new()
                {
                    TopLeftAddress = TopLeftAddress
                };
                ExcelOleObjectParameters oleSets = new()
                {
                    LinkToFile = false,
                    DisplayAsIcon = true
                };

                if (string.IsNullOrWhiteSpace(IconPath))
                {
                    oleSets.Icon = new ExcelImage(iconStream, ePictureType.Png);
                }
                else
                {
                    oleSets.Icon = new ExcelImage(IconPath);
                }
                ExcelOleObject oleObject = ws.Drawings.AddOleObject(Path.GetFileNameWithoutExtension(objectToEmbedPath), objectToEmbedPath, oleSets);
                oleObject.SetPosition(tmp.Row, IconX, tmp.Column, IcnoY);
                oleObject.SetSize(IconW, IconH);
                _logger.Info($"插入OLE对象到{ws.Name}完成");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "OLE对象插入失败");
            }
        }

        public static void ReadReportHeaderInfo(ExcelWorksheet ws, ReportHeaderViewModel reportHeaderInfo)
        {
            // 辅助函数: 找到issue和setup图片所在的标题行
            DataCell issueTitle = FindCellByValue(ws, "Issue Photos");
            DataCell setupTitle = FindCellByValue(ws, "Test Setup");

            reportHeaderInfo.TESTED_BY = FindInfoByText(ws, "TESTED BY");
            reportHeaderInfo.APPROVED_BY = FindInfoByText(ws, "APPROVED BY");
            reportHeaderInfo.PROJECT_NAME = FindInfoByText(ws, "PROJECT NAME");
            reportHeaderInfo.TEST_STAGE = FindInfoByText(ws, "TEST STAGE");
            reportHeaderInfo.TestDescription = FindInfoByText(ws, "Test Description");
            reportHeaderInfo.Test_Description_Pic = GetPicturesInRange(ws, 6, 1, 10);
            reportHeaderInfo.Issue_Photos_Pics = issueTitle is null ? null : GetPicturesInRange(ws, issueTitle.Row, 1, issueTitle.Row + 10);
            reportHeaderInfo.Test_Setup_Pics = setupTitle is null ? null : GetPicturesInRange(ws, setupTitle.Row, 1, setupTitle.Row + 10);
        }

        public static DataCell GetPicturesInRange(ExcelWorksheet ws, int startRow = 1, int startCol = 1, int endRow = -1, int endCol = -1)
        {
            if (ws == null || ws.Drawings.Count == 0)
            {
                return null;
            }

            var result = new DataCell()
            {
                Images = []
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
                            ImageSrc = Image.ConvertToWpfImage(picture.Image.ImageBytes),
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
            DataCell headerInfo = new();
            DataCell cell = FindCellByValue(ws, toFind);
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

        public static void ExcelAddPicture(ExcelWorksheet ws, string picName, DataCell pics, string TopLeft, string rpType, string tempPath)
        {
            if (pics.Images.Count <= 0)
            {
                return;
            }
            ExcelCellAddress start = new ExcelAddress(TopLeft).Start;
            int startRow = start.Row;
            int startCol = start.Column;
            for (int i = 0; i < pics.Images.Count; i++)
            {
                string picPath = Path.Combine(tempPath, picName + "_" + i + ".png");
                if (File.Exists(picPath))
                {
                    string[] temp = picPath.Split('.');
                    picPath = temp[0] + "_" + i + "." + temp[1];
                }

                Image.SaveImageSourceToFile(pics.Images[i].ImageSrc, picPath, "png");
                ExcelPicture test_desc_pic_excel = ws.Drawings.AddPicture(picName + "_" + i, picPath);
                test_desc_pic_excel.SetSize(300, 220);
                if (rpType.ToLower() == "burn")
                {
                    test_desc_pic_excel.SetPosition(startRow, 0, startCol + (i * 4), -18 + (i * 72));
                }
                else
                {
                    test_desc_pic_excel.SetPosition(startRow, 10, startCol + (i * 4), -24 + (i * 44));
                }
            }
        }


        #endregion

        #region 功能函数

        public static int ToInt(object obj, int defaultValue = 0)
        {
            return int.TryParse(obj?.ToString(), out var v) ? v : defaultValue;
        }

        public static string To_String(object obj, string defaultValue = "")
        {
            return obj?.ToString() ?? defaultValue;
        }

        public static string GetRelativePath(string relativeTo, string path)
        {
            // 1. 将路径转换为绝对路径并规范化（消除 . 和 .. 等）
            string fullPath = Path.GetFullPath(path);
            string fullRelativeTo = Path.GetFullPath(relativeTo);

            // 2. 确保基准路径以目录分隔符结尾，方便后续比较
            if (!fullRelativeTo.EndsWith(Path.DirectorySeparatorChar.ToString()))
            {
                fullRelativeTo += Path.DirectorySeparatorChar;
            }

            // 3. 检查是否共享同一个根目录（例如都在 C 盘）
            if (Path.GetPathRoot(fullPath) != Path.GetPathRoot(fullRelativeTo))
            {
                // 如果不在同一个盘符，无法计算相对路径，直接返回原绝对路径
                return fullPath;
            }

            // 4. 将路径按目录分隔符拆分
            var baseParts = fullRelativeTo.Split(new[] { Path.DirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries);
            var targetParts = fullPath.Split(new[] { Path.DirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries);

            // 5. 找出最长公共前缀的长度
            int commonLength = 0;
            int minLength = Math.Min(baseParts.Length, targetParts.Length);
            for (int i = 0; i < minLength; i++)
            {
                if (string.Equals(baseParts[i], targetParts[i], StringComparison.OrdinalIgnoreCase))
                {
                    commonLength++;
                }
                else
                {
                    break;
                }
            }

            // 6. 拼接 "../" 和剩余的相对路径
            var relativeParts = new System.Collections.Generic.List<string>();

            // 从基准路径向上回溯
            for (int i = commonLength; i < baseParts.Length; i++)
            {
                relativeParts.Add("..");
            }

            // 拼接目标路径多出来的部分
            for (int i = commonLength; i < targetParts.Length; i++)
            {
                relativeParts.Add(targetParts[i]);
            }

            return string.Join(Path.DirectorySeparatorChar.ToString(), relativeParts);
        }

        public static string GetTemplatePath(string rootPath, string reportType)
        {
            string[] excelExtensions = [".xlsx", ".xls", ".xlsm"];
            string[] excelFiles = Directory.GetFiles(rootPath, "*.*", SearchOption.AllDirectories).Where(file => excelExtensions.Contains(Path.GetExtension(file))).ToArray();
            Regex regex = new(@"[^a-zA-Z0-9]");
            foreach (string excelFile in excelFiles)
            {
                if (regex.Replace(Path.GetFileName(excelFile), "").ToLower().Contains(regex.Replace(reportType, "").ToLower()))
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

        #endregion
    }
}

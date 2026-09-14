using NLog;
using NPOI.SS.UserModel;
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


        #region Excel（NPOI）

        public static string GetCellAddress(int row, int column)
        {
            return ExcelNpoi.AddressOf(row, column);
        }
        public static string GetCellColumn(int column)
        {
            return ExcelNpoi.AddressOf(1, column).Replace("1", "");
        }

        public static DataCell FindCellByValue(ISheet ws, string value, string excludeValue = "", bool ignoreCase = true, DataCell start = null, DataCell end = null)
        {
            int snRowStart = 1;
            int snColumnStart = 1;
            int snColumnEnd = ExcelNpoi.LastColumn(ws);
            int snRowEnd = ExcelNpoi.LastRow(ws);
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
                    var _value = ExcelNpoi.CellText(ws, row, col);
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

        /* OLE 对象嵌入已迁移：NPOI 2.7.4 无 OLE 写入能力，改由 Utils/ExcelOleEmbedder.cs
           用 Excel COM 批量实现（流程变为：NPOI 写完并保存 → 调用 ExcelOleEmbedder.Embed）。 */


        public static void ReadReportHeaderInfo(ISheet ws, ReportHeaderViewModel reportHeaderInfo)
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

        public static DataCell GetPicturesInRange(ISheet ws, int startRow = 1, int startCol = 1, int endRow = -1, int endCol = -1)
        {
            if (ws == null || ws.DrawingPatriarch == null)
            {
                return null;
            }

            var result = new DataCell()
            {
                Images = []
            };

            List<(int Row, int Column, string Name, byte[] Bytes)> pictures = ExcelNpoi.Pictures(ws);
            if (endRow == -1)
            {
                endRow = ExcelNpoi.LastRow(ws);
            }
            if (endCol == -1)
            {
                endCol = ExcelNpoi.LastColumn(ws);
            }

            // 规范化范围 (防止用户传反了行列)
            int minRow = Math.Min(startRow, endRow);
            int maxRow = Math.Max(startRow, endRow);
            int minCol = Math.Min(startCol, endCol);
            int maxCol = Math.Max(startCol, endCol);

            foreach ((int Row, int Column, string Name, byte[] Bytes) picture in pictures)
            {
                // 图片左上角锚定的单元格坐标（1 基）
                int picRow = picture.Row;
                int picCol = picture.Column;

                // 判断逻辑：只要图片的左上角在指定范围内，就视为在该范围内
                if (picRow >= minRow && picRow <= maxRow &&
                    picCol >= minCol && picCol <= maxCol)
                {
                    result.Images.Add(new ExcelPictureInfo()
                    {
                        ImageSrc = Image.ConvertToWpfImage(picture.Bytes),
                        ImageBytes = picture.Bytes,
                        Name = picture.Name,
                    });
                    result.Data = "Images";
                    result.Row = picRow;
                    result.Column = picCol;
                }
            }
            result.Images.Reverse();
            return result;
        }

        public static DataCell FindInfoByText(ISheet ws, string toFind)
        {
            DataCell headerInfo = new();
            DataCell cell = FindCellByValue(ws, toFind);
            if (cell != null)
            {
                for (int c = cell.Column + 1; c <= ExcelNpoi.LastColumn(ws); c++)
                {
                    string value = ExcelNpoi.CellText(ws, cell.Row, c);
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

        public static void ExcelAddPicture(ISheet ws, string picName, DataCell pics, string TopLeft, string rpType, string tempPath)
        {
            if (pics?.Images == null || pics.Images.Count <= 0)
            {
                return;
            }
            int startRow = ExcelNpoi.RowOf(TopLeft);
            int startCol = ExcelNpoi.ColumnOf(TopLeft);
            if (startRow <= 0 || startCol <= 0)
            {
                _logger.Warn($"插入图片失败：单元格地址无效 {TopLeft}");
                return;
            }

            for (int i = 0; i < pics.Images.Count; i++)
            {
                ExcelPictureInfo info = pics.Images[i];
                byte[] bytes = info.ImageBytes;
                if (bytes == null || bytes.Length == 0)
                {
                    // 只有界面用 ImageSource 时，先落成临时 PNG 再读字节
                    string picPath = Path.Combine(tempPath, picName + "_" + i + ".png");
                    Image.SaveImageSourceToFile(info.ImageSrc, picPath, "png");
                    bytes = File.Exists(picPath) ? File.ReadAllBytes(picPath) : null;
                }
                if (bytes == null || bytes.Length == 0)
                {
                    _logger.Warn($"插入图片跳过：{picName}_{i} 无图片数据");
                    continue;
                }
                // 与原 EPPlus 版一致：300x220 像素，按序号横向每 4 列排一张
                int offsetY = rpType.ToLower() == "burn" ? -18 + (i * 72) : -24 + (i * 44);
                ExcelNpoi.AddPicture(ws.Workbook, ws, bytes, PictureType.PNG,
                    startRow, startCol + (i * 4), 300, 220, 0, Math.Max(0, offsetY));
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
            if (string.IsNullOrWhiteSpace(rootPath) || string.IsNullOrWhiteSpace(reportType))
            {
                return "";
            }
            if (!Directory.Exists(rootPath))
            {
                return "";
            }
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

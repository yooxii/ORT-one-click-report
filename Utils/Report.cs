using NLog;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using ORT一键报告.Models;
using ORT一键报告.Reports.Models;
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

            // 已完成的本地报告里还带有"测试周期/测试结论"，一并读取（模板里通常是空的，读不到就不覆盖界面既有值）
            DateTime? period = ReadTestPeriod(ws);
            if (period != null)
            {
                reportHeaderInfo.TestStart = period;
            }
            bool? conclusion = ReadTestConclusion(ws);
            if (conclusion != null)
            {
                reportHeaderInfo.TestPass = conclusion.Value;
            }
        }

        /// <summary>
        /// 读取报告里的"TEST PERIOD"起始日期（找不到或解析失败返回 null）
        /// </summary>
        public static DateTime? ReadTestPeriod(ISheet ws)
            => ParseReportDate(FindInfoByText(ws, "TEST PERIOD")?.Data);

        /// <summary>
        /// 读取报告里的"TEST CONCLUSION"是否为 Pass（找不到返回 null，由调用方保留默认值）
        /// </summary>
        public static bool? ReadTestConclusion(ISheet ws)
        {
            string value = FindInfoByText(ws, "TEST CONCLUSION")?.Data?.Trim();
            if (string.IsNullOrEmpty(value))
            {
                return null;
            }
            return value.StartsWith("pass", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 解析报告里的日期文本：支持 2025/7/2、2025-07-02、2025.7.2、2025年7月2日 等
        /// </summary>
        private static DateTime? ParseReportDate(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }
            Match m = Regex.Match(text, @"(\d{4})\s*[年/\-.]\s*(\d{1,2})\s*[月/\-.]\s*(\d{1,2})");
            if (m.Success
                && int.TryParse(m.Groups[1].Value, out int year)
                && int.TryParse(m.Groups[2].Value, out int month)
                && int.TryParse(m.Groups[3].Value, out int day))
            {
                try
                {
                    return new DateTime(year, month, day);
                }
                catch
                {
                    return null;
                }
            }
            // 只有月/日时按当前年份
            m = Regex.Match(text, @"(\d{1,2})\s*[月/\-.]\s*(\d{1,2})");
            if (m.Success
                && int.TryParse(m.Groups[1].Value, out int m2)
                && int.TryParse(m.Groups[2].Value, out int d2))
            {
                try
                {
                    return new DateTime(DateTime.Now.Year, m2, d2);
                }
                catch
                {
                    return null;
                }
            }
            return null;
        }

        /// <summary>
        /// 从报告文件读取表头信息并转成模型（供模型化生成使用）；
        /// 报告里没有的字段保持调用方给的默认值。找不到文件/标签返回 null。
        /// </summary>
        public static ReportHeaderData ReadHeaderData(string reportFilePath, int testTimeDays, ReportHeaderData fallback = null)
        {
            if (string.IsNullOrWhiteSpace(reportFilePath) || !File.Exists(reportFilePath))
            {
                return null;
            }
            // 报告文件也可能是 .xls，按内容选引擎
            IWorkbook wb = ExcelNpoi.OpenAny(reportFilePath);
            try
            {
                ISheet ws = ExcelNpoi.SheetAt(wb, 0);                ReportHeaderViewModel vm = new();
                ReadReportHeaderInfo(ws, vm);
                if (vm.TESTED_BY?.Data == null && vm.APPROVED_BY?.Data == null && vm.PROJECT_NAME?.Data == null)
                {
                    return null; // 报告里没有可用表头（例如拿到的是空模板）
                }
                DateTime start = vm.TestStart ?? fallback?.TestStart ?? DateTime.Now;
                ReportHeaderData header = new()
                {
                    TestedBy = FirstNonEmpty(vm.TESTED_BY?.Data, fallback?.TestedBy),
                    ApprovedBy = FirstNonEmpty(vm.APPROVED_BY?.Data, fallback?.ApprovedBy),
                    ProjectName = FirstNonEmpty(vm.PROJECT_NAME?.Data, fallback?.ProjectName),
                    TestStage = FirstNonEmpty(vm.TEST_STAGE?.Data, fallback?.TestStage),
                    TestDescription = FirstNonEmpty(vm.TestDescription?.Data, fallback?.TestDescription),
                    TestStart = start,
                    TestEnd = testTimeDays > 0 ? start.AddDays(testTimeDays) : (fallback?.TestEnd ?? start),
                    TestPass = vm.TestPass,
                    IssuePhotos = ToAttachments(vm.Issue_Photos_Pics),
                    TestSetupPhotos = ToAttachments(vm.Test_Setup_Pics),
                    TestDescriptionPhoto = ToAttachments(vm.Test_Description_Pic).FirstOrDefault()
                };
                return header;
            }
            catch (Exception ex)
            {
                _logger.Warn($"读取报告表头失败（{reportFilePath}）：{ex.Message}");
                return null;
            }
            finally
            {
                wb.Close();
            }
        }

        private static string FirstNonEmpty(string primary, string fallback)
            => string.IsNullOrWhiteSpace(primary) ? fallback : primary.Trim();

        /// <summary>
        /// 把图片单元格转成模型里的附件（保留字节，界面与生成共用）
        /// </summary>
        private static List<ImageAttachment> ToAttachments(DataCell cell)
        {
            List<ImageAttachment> list = [];
            if (cell?.Images == null)
            {
                return list;
            }
            foreach (ExcelPictureInfo pic in cell.Images)
            {
                list.Add(new ImageAttachment
                {
                    Name = pic.Name,
                    Bytes = pic.ImageBytes
                });
            }
            return list;
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
                // 按文件头识别格式（EPPlus 是自动识别的）；EMF/WMF/TIFF 先转 PNG，单张失败只跳过该图
                PictureType pictureType = ExcelNpoi.DetectPictureType(bytes);
                if (pictureType is PictureType.EMF or PictureType.WMF or PictureType.TIFF)
                {
                    byte[] png = ExcelNpoi.TryConvertToPng(bytes);
                    if (png == null)
                    {
                        _logger.Warn($"插入图片跳过：{picName}_{i} 格式 {pictureType} 转换 PNG 失败");
                        continue;
                    }
                    bytes = png;
                    pictureType = PictureType.PNG;
                }
                // 与原 EPPlus 版一致：300x220 像素，按序号横向每 4 列排一张
                int offsetY = rpType.ToLower() == "burn" ? -18 + (i * 72) : -24 + (i * 44);
                try
                {
                    ExcelNpoi.AddPicture(ws.Workbook, ws, bytes, pictureType,
                        startRow, startCol + (i * 4), 300, 220, 0, Math.Max(0, offsetY));
                }
                catch (Exception ex)
                {
                    _logger.Warn($"插入图片失败（跳过 {picName}_{i}，格式 {pictureType}）：{ex.Message}");
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

        /// <summary>
        /// 从文件夹/文件名里取出 WK#### 原样文本（如 "WK2525"）；取不到返回 null
        /// </summary>
        public static string ParseWeekTag(string pathOrName)
        {
            Match match = Regex.Match(pathOrName ?? "", @"WK\s*(\d{4})", RegexOptions.IgnoreCase);
            return match.Success ? "WK" + match.Groups[1].Value : null;
        }

        /// <summary>
        /// 从报告文件夹/文件名里的 WK#### 解析测试周期（ISO 周：WK2506 = 2025 年第 6 周），返回该周周一。
        /// 解析失败返回 null（.NET Framework 4.8 没有 ISOWeek，这里自己算）。
        /// </summary>
        public static DateTime? ParseWeekPeriod(string pathOrName)
        {
            Match match = Regex.Match(pathOrName ?? "", @"WK\s*(\d{2})(\d{2})", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                return null;
            }
            int year = 2000 + int.Parse(match.Groups[1].Value);
            int week = int.Parse(match.Groups[2].Value);
            if (week < 1 || week > 53)
            {
                return null;
            }
            try
            {
                // ISO 周：含 1 月 4 日的那一周是第 1 周，周一为一周开始
                DateTime jan4 = new(year, 1, 4);
                int offsetToMonday = ((int)jan4.DayOfWeek + 6) % 7;
                DateTime firstMonday = jan4.AddDays(-offsetToMonday);
                return firstMonday.AddDays((week - 1) * 7);
            }
            catch
            {
                return null;
            }
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

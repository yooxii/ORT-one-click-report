using Microsoft.Office.Interop.Word;
using NLog;
using System;
using System.IO;
using System.Linq;

namespace ORT一键报告.Utils
{
    public class Docx2Pdf
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// 转换单个 Word 文档为 PDF
        /// </summary>
        public static void ConvertToPdf(string sourcePath, string targetPath)
        {
            Application wordApp = new();
            try
            {
                wordApp.Visible = false;
                ConvertSingleFile(wordApp, sourcePath, targetPath);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "转换过程中发生错误: ");
            }
            finally
            {
                wordApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            }
        }

        /// <summary>
        /// 转换指定目录下所有 docx 文件为 PDF
        /// </summary>
        public static void ConvertToPdf(string sourceDir)
        {
            Application wordApp = new();
            wordApp.Visible = false;
            _logger.Info("{0} 转换PDF开始", sourceDir);
            try
            {
                if (!Directory.Exists(sourceDir))
                {
                    throw new DirectoryNotFoundException($"{sourceDir}不存在");
                }

                var files = Directory.GetFiles(sourceDir).Where(f => Path.GetExtension(f).Contains("docx"));

                foreach (var file in files)
                {
                    string targetPath = file.Replace("docx", "pdf");

                    if (File.Exists(targetPath))
                    {
                        continue;
                    }
                    ConvertSingleFile(wordApp, file, targetPath);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "转换过程中发生错误: ");
            }
            finally
            {
                wordApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            }
        }

        /// <summary>
        /// 使用已创建的 Word 应用程序实例转换单个文件
        /// </summary>
        private static void ConvertSingleFile(Application wordApp, string sourcePath, string targetPath)
        {
            Document wordDoc = null;
            try
            {
                wordDoc = wordApp.Documents.Open(sourcePath);
                wordDoc.ExportAsFixedFormat(targetPath, WdExportFormat.wdExportFormatPDF);
                _logger.Info("转换成功！PDF 已保存至: {0}", targetPath);
            }
            finally
            {
                if (wordDoc != null)
                {
                    wordDoc.Close(WdSaveOptions.wdDoNotSaveChanges);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc);
                }
            }
        }

        public static void AlertFileTime(string sourceDir)
        {
            try
            {
                if (!Directory.Exists(sourceDir))
                {
                    throw new DirectoryNotFoundException($"{sourceDir}不存在");
                }

                var files = Directory.GetFiles(sourceDir).Where(f => Path.GetExtension(f).Contains("docx"));

                var startTime = File.GetLastWriteTime(files.First()).AddDays(new Random().Next(7));
                if (startTime.DayOfWeek > DayOfWeek.Friday)
                {
                    startTime.AddDays(new Random().Next(2, 5));
                }

                foreach (var file in files)
                {
                    string targetPath = file.Replace("docx", "pdf");
                    var nowTime = File.GetLastWriteTime(targetPath);
                    var newTime = new DateTime(startTime.Year, startTime.Month, startTime.Day, nowTime.Hour, nowTime.Minute, nowTime.Second);
                    File.SetCreationTime(targetPath, newTime);
                    File.SetLastWriteTime(targetPath, newTime);
                    _logger.Info("{0} : {1}", targetPath, File.GetLastWriteTime(targetPath));
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex);
            }
        }
    }
}

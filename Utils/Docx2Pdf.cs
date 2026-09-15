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
        /// 转换单个 Word 文档为 PDF；返回是否成功（以产物存在且非空为准，不再静默失败）
        /// </summary>
        public static bool ConvertToPdf(string sourcePath, string targetPath)
        {
            if (string.IsNullOrWhiteSpace(sourcePath) || string.IsNullOrWhiteSpace(targetPath))
            {
                return false;
            }
            Application wordApp = new();
            try
            {
                wordApp.Visible = false;
                ConvertSingleFile(wordApp, sourcePath, targetPath);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"转换 PDF 失败: {sourcePath}");
            }
            finally
            {
                try
                {
                    wordApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                }
                catch
                {
                    // Word 可能已退出，忽略
                }
            }
            return File.Exists(targetPath) && new FileInfo(targetPath).Length > 0;
        }

        /// <summary>
        /// 转换指定目录下所有 docx 文件为 PDF（复用同一个 Word 实例）。
        /// 关键点：逐个文件独立 try/catch —— 单个文件失败不会中断整批（原实现是一层 try 包整个循环，
        /// 一个文件打不开就会导致后面的 Word 全部不转换）；返回(成功数, 失败数)。
        /// </summary>
        public static (int Converted, int Failed) ConvertToPdf(string sourceDir)
        {
            int converted = 0, failed = 0;
            if (!Directory.Exists(sourceDir))
            {
                _logger.Error($"{sourceDir}不存在");
                return (0, 0);
            }
            Application wordApp = new();
            wordApp.Visible = false;
            _logger.Info("{0} 转换PDF开始", sourceDir);
            try
            {
                var files = Directory.GetFiles(sourceDir).Where(f => Path.GetExtension(f).Contains("docx"));

                foreach (var file in files)
                {
                    string targetPath = Path.ChangeExtension(file, ".pdf");

                    if (File.Exists(targetPath) && new FileInfo(targetPath).Length > 0)
                    {
                        continue;
                    }
                    try
                    {
                        ConvertSingleFile(wordApp, file, targetPath);
                        converted++;
                    }
                    catch (Exception ex)
                    {
                        failed++;
                        _logger.Error(ex, $"转换 PDF 失败（继续处理后续文件）: {file}");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "转换过程中发生错误: ");
            }
            finally
            {
                try
                {
                    wordApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                }
                catch
                {
                    // Word 可能已退出，忽略
                }
            }
            _logger.Info($"转换PDF结束：成功 {converted} 个，失败 {failed} 个（{sourceDir}）");
            return (converted, failed);
        }

        /// <summary>
        /// 使用已创建的 Word 应用程序实例转换单个文件。
        /// 打开文档时把参数写全并显式关掉确认转换/最近文件/可见性：
        /// 只传文件名时 Word 可能想弹确认框或走"受保护视图"，自动化下会失败并抛
        /// "Word 未能引发事件 (0x800A1772)"，这是该错误最常见的成因。
        /// </summary>
        private static void ConvertSingleFile(Application wordApp, string sourcePath, string targetPath)
        {
            Document wordDoc = null;
            try
            {
                wordDoc = OpenDocument(wordApp, sourcePath);
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

        /// <summary>
        /// 打开 Word 文档；若直接打开失败（受保护视图/来源受限），先复制到本地临时目录再试一次
        /// </summary>
        private static Document OpenDocument(Application wordApp, string sourcePath)
        {
            try
            {
                return wordApp.Documents.Open(
                    FileName: sourcePath,
                    ConfirmConversions: false,
                    ReadOnly: true,
                    AddToRecentFiles: false,
                    Visible: false,
                    OpenAndRepair: false);
            }
            catch (Exception ex)
            {
                _logger.Warn($"直接打开失败({ex.Message})，复制到临时目录后重试：{sourcePath}");
                string localCopy = Path.Combine(Path.GetTempPath(), "ort_docx_" + Guid.NewGuid().ToString("N") + Path.GetExtension(sourcePath));
                File.Copy(sourcePath, localCopy, true);
                try
                {
                    return wordApp.Documents.Open(
                        FileName: localCopy,
                        ConfirmConversions: false,
                        ReadOnly: true,
                        AddToRecentFiles: false,
                        Visible: false,
                        OpenAndRepair: false);
                }
                finally
                {
                    // 文档已被 Word 打开后再删除会失败，交给系统临时目录清理
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

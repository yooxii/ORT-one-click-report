using Microsoft.Office.Interop.Word;
using NLog;
using System;
using System.IO;

namespace ORT一键报告
{
    public class Docx2Pdf
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();
        public static void ConvertToPdf(string sourcePath, string targetPath = "")
        {
            // 1. 创建 Word 应用程序实例
            Application wordApp = new Application();
            Document wordDoc = null;
            if (targetPath == "")
            {
                targetPath = Path.GetDirectoryName(sourcePath) + Path.GetFileNameWithoutExtension(sourcePath) + "pdf";
            }
            try
            {
                // 2. 设置为不可见模式（后台静默运行）
                wordApp.Visible = false;

                // 3. 打开指定的 Word 文档
                wordDoc = wordApp.Documents.Open(sourcePath);

                // 4. 调用导出功能，将文档保存为 PDF 格式
                // WdExportFormat.wdExportFormatPDF 是导出为 PDF 的枚举常量
                wordDoc.ExportAsFixedFormat(targetPath, WdExportFormat.wdExportFormatPDF);

                _logger.Info("转换成功！PDF 已保存至: " + targetPath);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "转换过程中发生错误: " + ex.Message);
            }
            finally
            {
                // 5. 无论成功与否，都要确保关闭文档和退出 Word 程序
                if (wordDoc != null)
                {
                    // 关闭文档，不保存对原文档的修改
                    wordDoc.Close(WdSaveOptions.wdDoNotSaveChanges);
                }
                wordApp.Quit();

                // 6. 释放 COM 对象，防止后台残留 WINWORD.EXE 进程
                if (wordDoc != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            }
        }
    }
}

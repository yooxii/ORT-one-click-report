using NLog;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace ORT一键报告.Utils
{
    /// <summary>
    /// 一个待嵌入的 OLE 对象（SN/ATE 等附件）
    /// </summary>
    public class OleEmbedRequest
    {
        /// <summary>要嵌入的文件完整路径</summary>
        public string ObjectPath { get; set; }

        /// <summary>嵌入到的工作表名（空则用第一个工作表）</summary>
        public string SheetName { get; set; }

        /// <summary>左上角单元格地址，如 "F12"</summary>
        public string TopLeftAddress { get; set; }

        /// <summary>自定义图标路径（EMF/PNG/图标文件），空则用内置的 Excel 图标</summary>
        public string IconPath { get; set; }

        /// <summary>图标标签文字</summary>
        public string IconLabel { get; set; } = "点击查看详细数据";

        /// <summary>相对锚点单元格的偏移（像素）</summary>
        public int OffsetXPx { get; set; }
        public int OffsetYPx { get; set; }

        /// <summary>图标尺寸（像素，&lt;=0 表示用 Excel 默认）</summary>
        public int WidthPx { get; set; }
        public int HeightPx { get; set; }
    }

    /// <summary>
    /// OLE 对象嵌入：NPOI 2.7.4 没有 OLE 写入能力，改用 Excel COM（Interop）实现。
    /// 说明：批量接口只开一次 Excel 会话，避免逐条开关 Excel 造成的性能问题。
    /// 需要目标机器安装 Excel（本程序 ATE 流程本来也依赖 Excel COM）。
    /// </summary>
    public static class ExcelOleEmbedder
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        /// <summary>Excel COM 里 1 像素对应的磅值（96 DPI → 72 磅/英寸）</summary>
        private const double PointsPerPixel = 72.0 / 96.0;

        /// <summary>
        /// 把若干对象嵌入到指定 xlsx（文件会就地保存）
        /// </summary>
        public static void Embed(string xlsxPath, IReadOnlyList<OleEmbedRequest> requests)
        {
            List<OleEmbedRequest> items = requests?.Where(r => r != null && !string.IsNullOrWhiteSpace(r.ObjectPath)).ToList() ?? [];
            if (items.Count == 0)
            {
                return;
            }
            if (!File.Exists(xlsxPath))
            {
                _logger.Warn($"OLE 嵌入跳过：文件不存在 {xlsxPath}");
                return;
            }

            string iconFile = null;
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Microsoft.Office.Interop.Excel.Workbook workbook = null;
            try
            {
                iconFile = EnsureIconFile(items[0].IconPath);
                excelApp = new Microsoft.Office.Interop.Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false,
                    ScreenUpdating = false
                };
                workbook = excelApp.Workbooks.Open(xlsxPath);
                foreach (OleEmbedRequest req in items)
                {
                    EmbedOne(workbook, req, iconFile);
                }
                workbook.Save();
                _logger.Info($"已嵌入 {items.Count} 个 OLE 对象到 {Path.GetFileName(xlsxPath)}");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"OLE 对象嵌入失败：{xlsxPath}");
            }
            finally
            {
                if (workbook != null)
                {
                    try
                    {
                        workbook.Close(SaveChanges: false);
                    }
                    catch (Exception ex)
                    {
                        _logger.Warn($"关闭工作簿失败（忽略）：{ex.Message}");
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    try
                    {
                        excelApp.Quit();
                    }
                    catch (Exception ex)
                    {
                        _logger.Warn($"退出 Excel 失败（忽略）：{ex.Message}");
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                TryDelete(iconFile);
            }
        }

        /// <summary>
        /// 嵌入单个对象（内部实现）
        /// </summary>
        private static void EmbedOne(Microsoft.Office.Interop.Excel.Workbook workbook, OleEmbedRequest req, string iconFile)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = string.IsNullOrWhiteSpace(req.SheetName)
                ? (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1]
                : (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[req.SheetName];
            Microsoft.Office.Interop.Excel.Range range = sheet.Range[req.TopLeftAddress];
            double left = (double)range.Left + req.OffsetXPx * PointsPerPixel;
            double top = (double)range.Top + req.OffsetYPx * PointsPerPixel;

            // OLEObjects() 的返回类型是 object（COM 晚期绑定），用 dynamic 才能调用 Add
            dynamic oleObjects = sheet.OLEObjects();
            dynamic ole = oleObjects.Add(
                Filename: req.ObjectPath,
                Link: false,
                DisplayAsIcon: true,
                IconFileName: (object)iconFile ?? Type.Missing,
                IconIndex: Type.Missing,
                IconLabel: req.IconLabel,
                Left: left,
                Top: top);
            if (req.WidthPx > 0)
            {
                ole.Width = req.WidthPx * PointsPerPixel;
            }
            if (req.HeightPx > 0)
            {
                ole.Height = req.HeightPx * PointsPerPixel;
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ole);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
        }

        /// <summary>
        /// 准备 OLE 图标文件：Excel 的 IconFileName 只接受 .ico/.exe/.dll，
        /// 这里把内置的 EMF（或调用方给的图片）转成临时 .ico；失败返回 null（改用 Excel 默认图标）。
        /// </summary>
        private static string EnsureIconFile(string iconPath)
        {
            string tempFile = null;
            try
            {
                byte[] bytes;
                if (string.IsNullOrWhiteSpace(iconPath))
                {
                    bytes = Resources.image_xlsx_emf;
                }
                else if (File.Exists(iconPath))
                {
                    bytes = File.ReadAllBytes(iconPath);
                }
                else
                {
                    _logger.Warn($"OLE 图标文件不存在，改用内置图标：{iconPath}");
                    bytes = Resources.image_xlsx_emf;
                }
                if (bytes == null || bytes.Length == 0)
                {
                    return null;
                }

                using MemoryStream stream = new(bytes);
                using System.Drawing.Image source = System.Drawing.Image.FromStream(stream);
                using Bitmap bitmap = new(source, new Size(32, 32));
                IntPtr handle = bitmap.GetHicon();
                using Icon icon = Icon.FromHandle(handle);
                tempFile = Path.Combine(Path.GetTempPath(), "ort_ole_icon_" + Guid.NewGuid().ToString("N") + ".ico");
                using (FileStream fs = new(tempFile, FileMode.Create, FileAccess.Write))
                {
                    icon.Save(fs);
                }
                DestroyIcon(handle);
                return tempFile;
            }
            catch (Exception ex)
            {
                _logger.Warn($"生成 OLE 图标失败，改用 Excel 默认图标：{ex.Message}");
                TryDelete(tempFile);
                return null;
            }
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
                // 临时文件删不掉不影响功能
            }
        }

        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        private static extern bool DestroyIcon(IntPtr handle);
    }
}

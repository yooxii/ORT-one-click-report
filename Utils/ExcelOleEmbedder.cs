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
    /// OLE 嵌入结果（用于把"是否真的写进文件"如实反馈给调用方/界面）
    /// </summary>
    public class OleEmbedResult
    {
        /// <summary>请求嵌入的对象数</summary>
        public int Requested { get; set; }

        /// <summary>Excel 接受并添加成功的对象数</summary>
        public int Added { get; set; }

        /// <summary>添加失败（跳过）的对象数</summary>
        public int Failed { get; set; }

        /// <summary>Excel 会话崩溃重启次数</summary>
        public int Restarts { get; set; }

        /// <summary>是否确认已保存到文件（false 表示附件可能没写进去）</summary>
        public bool Saved { get; set; }

        /// <summary>是否因为没装 Excel 而整体跳过</summary>
        public bool SkippedNoExcel { get; set; }

        /// <summary>是否由内置直写方式（不依赖 Excel）完成嵌入</summary>
        public bool UsedDirectWriter { get; set; }

        /// <summary>给用户看的一句话结论</summary>
        public string Summary =>
            Saved && UsedDirectWriter ? $"附件已嵌入 {Added}/{Requested} 个（Excel 不可用，已由内置方式写入）"
            : !Saved ? $"附件嵌入未确认落盘（已尝试 {Added}/{Requested} 个，详见日志）"
            : Failed > 0 ? $"附件已嵌入 {Added}/{Requested} 个，{Failed} 个失败（详见日志）"
            : $"附件已嵌入 {Added}/{Requested} 个";
    }

    /// <summary>
    /// OLE 对象嵌入：优先用 Excel COM（后期绑定，不依赖 Interop PIA），
    /// 若未安装 Excel 或 Excel 没能把附件写进文件，则自动退回 Utils/ExcelOleWriter 直写包的方式，
    /// 保证"附件嵌入"这一功能不因 Excel 环境问题而丢失（原 EPPlus 版本同样不依赖 Excel）。
    /// </summary>
    public static class ExcelOleEmbedder
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        /// <summary>Excel COM 里 1 像素对应的磅值（96 DPI → 72 磅/英寸）</summary>
        private const double PointsPerPixel = 72.0 / 96.0;

        /// <summary>每嵌入多少个对象保存一次（对象很多时 Excel 可能崩溃，分批保存可保住已完成的部分）</summary>
        private const int BatchSize = 5;

        /// <summary>Excel 会话崩溃后的最大重启次数</summary>
        private const int MaxRestarts = 15;

        /// <summary>
        /// 把若干对象嵌入到指定 xlsx（文件会就地保存）。
        /// 说明：Excel 在连续嵌入大量 OLE 对象时可能自身崩溃（RPC 服务器不可用），
        /// 因此这里分批保存，并在检测到会话中断时重启 Excel 并从断点继续，避免整批失败。
        /// </summary>
        public static OleEmbedResult Embed(string xlsxPath, IReadOnlyList<OleEmbedRequest> requests)
        {
            OleEmbedResult result = new() { Requested = requests?.Count(r => r != null && !string.IsNullOrWhiteSpace(r.ObjectPath)) ?? 0 };
            List<OleEmbedRequest> items = requests?.Where(r => r != null && !string.IsNullOrWhiteSpace(r.ObjectPath)).ToList() ?? [];
            if (items.Count == 0)
            {
                result.Saved = true;
                return result;
            }
            if (!File.Exists(xlsxPath))
            {
                _logger.Warn($"OLE 嵌入跳过：文件不存在 {xlsxPath}");
                return result;
            }
            // 后期绑定调用 Excel COM：不依赖 Interop PIA / office.dll，只要求装了 Excel
            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                _logger.Warn("未检测到 Excel，改用内置直写方式嵌入 OLE 附件");
                result.SkippedNoExcel = true;
                UseDirectWriter(xlsxPath, items, result);
                return result;
            }

            string iconFile = null;
            dynamic excelApp = null;
            dynamic workbook = null;
            int index = 0, embedded = 0, failed = 0, restarts = 0, sinceSave = 0;
            bool saved = false;
            try
            {
                iconFile = EnsureIconFile(items[0].IconPath);
                while (index < items.Count && restarts <= MaxRestarts)
                {
                    if (excelApp == null)
                    {
                        excelApp = Activator.CreateInstance(excelType);
                        excelApp.Visible = false;
                        excelApp.DisplayAlerts = false;
                        excelApp.ScreenUpdating = false;
                        excelApp.AskToUpdateLinks = false;
                        workbook = OpenWorkbook(excelApp, xlsxPath);
                        sinceSave = 0;
                    }

                    bool crashed = false;
                    OleEmbedRequest req = items[index];
                    try
                    {
                        EmbedOne(workbook, req, iconFile);
                        embedded++;
                        index++;
                        sinceSave++;
                    }
                    catch (Exception ex)
                    {
                        if (IsSessionLost(ex))
                        {
                            crashed = true;
                        }
                        else
                        {
                            failed++;
                            index++;
                            sinceSave++;
                            _logger.Warn($"嵌入 OLE 失败（跳过该条）：{Path.GetFileName(req.ObjectPath)} - {ex.Message}");
                        }
                    }

                    if (!crashed && sinceSave >= BatchSize)
                    {
                        if (TrySave(workbook, xlsxPath))
                        {
                            sinceSave = 0;
                            saved = true;
                        }
                    }

                    if (crashed)
                    {
                        restarts++;
                        _logger.Warn($"Excel 会话中断（第 {restarts} 次），已嵌入 {embedded}/{items.Count}，重启后继续");
                        CloseExcel(ref excelApp, ref workbook);
                        System.Threading.Thread.Sleep(1500);
                    }
                }

                if (workbook != null)
                {
                    if (TrySave(workbook, xlsxPath))
                    {
                        saved = true;
                    }
                    else
                    {
                        _logger.Warn($"最终保存未确认落盘：{Path.GetFileName(xlsxPath)}");
                    }
                }
                _logger.Info($"OLE 嵌入完成：成功 {embedded}，失败 {failed}，Excel 重启 {restarts} 次，共 {items.Count} 个，落盘={saved}（{Path.GetFileName(xlsxPath)}）");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"OLE 对象嵌入失败：{xlsxPath}");
            }
            finally
            {
                CloseExcel(ref excelApp, ref workbook);
                TryDelete(iconFile);
            }

            result.Added = embedded;
            result.Failed = failed;
            result.Restarts = restarts;
            result.Saved = saved && index >= items.Count;

            // Excel 没能把附件写进文件（未保存/会话崩溃导致中断）时，退回直写包方式兜底
            if (!result.Saved)
            {
                UseDirectWriter(xlsxPath, items, result);
            }
            return result;
        }

        /// <summary>
        /// 直写包兜底：不依赖 Excel，直接把附件写进 xlsx
        /// </summary>
        private static void UseDirectWriter(string xlsxPath, IReadOnlyList<OleEmbedRequest> items, OleEmbedResult result)
        {
            try
            {
                int direct = ExcelOleWriter.Write(xlsxPath, items);
                if (direct > 0)
                {
                    result.UsedDirectWriter = true;
                    result.Added = direct;
                    result.Failed = Math.Max(0, items.Count - direct);
                    result.Saved = true;
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"直写 OLE 附件失败：{xlsxPath}");
            }
        }

        /// <summary>
        /// 判断异常是否为 Excel 会话中断（进程崩溃/断开），这类错误需要重启 Excel 才能继续
        /// </summary>
        private static bool IsSessionLost(Exception ex)
        {
            int hr = ex.HResult;
            return hr == unchecked((int)0x800706BA)   // RPC 服务器不可用
                || hr == unchecked((int)0x800706BE)   // RPC 调用失败
                || hr == unchecked((int)0x80010108)   // RPC_E_DISCONNECTED
                || hr == unchecked((int)0x800401FD)   // 对象未连接到服务器
                || hr == unchecked((int)0x80010007)   // RPC_E_SERVER_DIED
                || (ex.Message?.Contains("RPC") ?? false)
                || (ex.Message?.Contains("远程过程调用") ?? false);
        }

        /// <summary>
        /// 打开工作簿并确保是可写状态：
        /// Excel 崩溃后会残留锁文件（~$xxx.xlsx），下一个 Excel 会话会把文件当成"已被占用"而只读打开，
        /// 此时所有 Save 都会失败（报"文档未保存"），所以先清锁文件，只读时清理后重开一次。
        /// </summary>
        private static dynamic OpenWorkbook(dynamic excelApp, string xlsxPath)
        {
            ClearLockFile(xlsxPath);
            dynamic workbook = excelApp.Workbooks.Open(xlsxPath, 0, false);
            if ((bool)workbook.ReadOnly)
            {
                _logger.Warn("Excel 以只读方式打开了文件，清理锁文件后重开");
                try
                {
                    workbook.Close(false);
                }
                catch
                {
                    // 忽略
                }
                Release((object)workbook);
                ClearLockFile(xlsxPath);
                workbook = excelApp.Workbooks.Open(xlsxPath, 0, false);
            }
            return workbook;
        }

        /// <summary>
        /// 清理与目标文件同名的 Excel 残留锁文件（~$ 开头）
        /// </summary>
        private static void ClearLockFile(string xlsxPath)
        {
            try
            {
                string dir = Path.GetDirectoryName(xlsxPath);
                string stem = Path.GetFileNameWithoutExtension(xlsxPath);
                if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir))
                {
                    return;
                }
                foreach (string lockFile in Directory.GetFiles(dir, "~$*"))
                {
                    string lockStem = Path.GetFileNameWithoutExtension(lockFile);
                    lockStem = lockStem.Length > 2 ? lockStem.Substring(2) : lockStem;
                    // Excel 锁文件名会截断长文件名，这里双向前缀匹配
                    if (stem.StartsWith(lockStem, StringComparison.OrdinalIgnoreCase)
                        || lockStem.StartsWith(stem, StringComparison.OrdinalIgnoreCase))
                    {
                        File.Delete(lockFile);
                        _logger.Warn($"已清理 Excel 残留锁文件：{Path.GetFileName(lockFile)}");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"清理锁文件失败（忽略）：{ex.Message}");
            }
        }

        /// <summary>
        /// 保存并确认真的落盘（Save 失败时退回 SaveAs；最后用文件时间戳校验）
        /// </summary>
        private static bool TrySave(dynamic workbook, string xlsxPath)
        {
            DateTime before = File.Exists(xlsxPath) ? File.GetLastWriteTimeUtc(xlsxPath) : DateTime.MinValue;
            try
            {
                workbook.Save();
            }
            catch (Exception ex)
            {
                _logger.Warn($"Excel Save 失败：{ex.Message}");
                try
                {
                    workbook.SaveAs(xlsxPath, 51); // 51 = xlOpenXMLWorkbook
                }
                catch (Exception ex2)
                {
                    _logger.Warn($"Excel SaveAs 也失败：{ex2.Message}");
                    return false;
                }
            }
            System.Threading.Thread.Sleep(400);
            DateTime after = File.Exists(xlsxPath) ? File.GetLastWriteTimeUtc(xlsxPath) : DateTime.MinValue;
            if (after <= before)
            {
                _logger.Warn($"保存后文件时间戳未变化（可能未真正落盘）：{Path.GetFileName(xlsxPath)}");
                return false;
            }
            return true;
        }

        /// <summary>
        /// 关闭并释放 Excel 会话（进程可能已崩溃，失败一律忽略）
        /// </summary>
        private static void CloseExcel(ref dynamic excelApp, ref dynamic workbook)
        {
            try
            {
                if (workbook != null)
                {
                    try
                    {
                        workbook.Close(false);
                    }
                    catch
                    {
                        // 会话已断，无需处理
                    }
                    Release((object)workbook);
                }
            }
            catch
            {
                // dynamic 封送失败（会话已断）时忽略
            }
            workbook = null;

            try
            {
                if (excelApp != null)
                {
                    try
                    {
                        excelApp.Quit();
                    }
                    catch
                    {
                        // 会话已断，无需处理
                    }
                    Release((object)excelApp);
                }
            }
            catch
            {
                // dynamic 封送失败（会话已断）时忽略
            }
            excelApp = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// 嵌入单个对象（内部实现，后期绑定）
        /// </summary>
        private static void EmbedOne(dynamic workbook, OleEmbedRequest req, string iconFile)
        {
            dynamic sheet = string.IsNullOrWhiteSpace(req.SheetName)
                ? workbook.Worksheets[1]
                : workbook.Worksheets[req.SheetName];
            dynamic range = sheet.Range[req.TopLeftAddress];
            double left = (double)range.Left + req.OffsetXPx * PointsPerPixel;
            double top = (double)range.Top + req.OffsetYPx * PointsPerPixel;

            dynamic oleObjects = sheet.OLEObjects();
            // 必须用命名参数：OLEObjects.Add 的第一个参数是 ClassType，Filename 在第二位
            dynamic ole = oleObjects.Add(
                ClassType: Type.Missing,
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
            // 释放时对象可能已经失效（Excel 崩溃），这里必须整体兜住：dynamic 参数的封送本身也可能抛 COM 异常
            try
            {
                Release((object)ole);
                Release((object)range);
                Release((object)sheet);
            }
            catch
            {
                // 会话已断，释放失败无需处理
            }
        }

        /// <summary>
        /// 释放 COM 对象（先转成 object，避免 dynamic 封送在 try 之外抛异常）
        /// </summary>
        private static void Release(object comObject)
        {
            try
            {
                if (comObject != null && System.Runtime.InteropServices.Marshal.IsComObject(comObject))
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(comObject);
                }
            }
            catch
            {
                // 释放失败不影响功能
            }
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

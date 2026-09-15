using NLog;
using ORT一键报告.Services;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace ORT一键报告.Reports.Views
{
    /// <summary>
    /// EMI 数据处理：① 按用户参数设置 PDF 文件时间；② 批量替换 Word 文档里的字符串。
    /// 两个操作都作用于「EMI 测试数据文件夹」。
    /// </summary>
    public partial class WindowEmiTools : Window
    {
        private readonly IPathService _pathService;
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public WindowEmiTools(IPathService pathService, string dataDir)
        {
            InitializeComponent();
            _pathService = pathService;
            txt_dataDir.Text = dataDir ?? "";
            dp_baseDate.SelectedDate = DateTime.Today;
            txt_randomDays.Text = "0";
        }

        private string DataDir => txt_dataDir.Text?.Trim() ?? "";

        private void Btn_SelectDir_Click(object sender, RoutedEventArgs e)
        {
            string dir = _pathService.OpenPathDialog("选择EMI测试数据文件夹", initPath: DataDir, isDir: true);
            if (!string.IsNullOrWhiteSpace(dir))
            {
                txt_dataDir.Text = dir;
            }
        }

        private async void Btn_ApplyTime_Click(object sender, RoutedEventArgs e)
        {
            if (!Directory.Exists(DataDir))
            {
                SetStatus("数据目录不存在，请先选择。");
                return;
            }
            DateTime baseDate = dp_baseDate.SelectedDate ?? DateTime.Today;
            int randomDays = int.TryParse(txt_randomDays.Text?.Trim(), out int parsed) && parsed >= 0 ? parsed : 0;
            bool skipWeekend = chk_skipWeekend.IsChecked == true;

            SetBusy(true, "正在设置 PDF 文件时间…");
            try
            {
                int count = await Task.Run(() => Docx2Pdf.SetPdfFileTimes(DataDir, baseDate, randomDays, skipWeekend));
                SetStatus($"已设置 {count} 个 PDF 的文件时间（基准 {baseDate:yyyy-MM-dd}，随机 0~{randomDays} 天，周末顺延={(skipWeekend ? "是" : "否")}）。");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "设置 PDF 文件时间失败");
                SetStatus($"设置失败：{ex.Message}");
            }
            finally
            {
                SetBusy(false, null);
            }
        }

        private async void Btn_Replace_Click(object sender, RoutedEventArgs e)
        {
            List<KeyValuePair<string, string>> pairs = DocxTextReplacer.ParsePairs(txt_pairs.Text);
            if (pairs.Count == 0)
            {
                SetStatus("请按 旧=新 每行一组填写替换规则。");
                return;
            }
            if (!Directory.Exists(DataDir))
            {
                SetStatus("数据目录不存在，请先选择。");
                return;
            }
            bool backup = chk_backup.IsChecked == true;

            SetBusy(true, $"正在替换（{pairs.Count} 组规则）…");
            try
            {
                DocxReplaceResult result = await Task.Run(() => DocxTextReplacer.ReplaceInFolder(DataDir, pairs, backup));
                string message = $"替换完成：扫描 {result.Files} 个 Word 文件，改动 {result.ChangedFiles} 个，共替换 {result.Replacements} 处"
                    + (result.Failed > 0 ? $"，失败 {result.Failed} 个" : "") + "。";
                if (result.Messages.Count > 0)
                {
                    message += "\n" + string.Join("\n", result.Messages.Take(5));
                }
                SetStatus(message);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Word 字符串替换失败");
                SetStatus($"替换失败：{ex.Message}");
            }
            finally
            {
                SetBusy(false, null);
            }
        }

        private void Btn_Close_Click(object sender, RoutedEventArgs e) => Close();

        private void SetBusy(bool busy, string status)
        {
            btn_applyTime.IsEnabled = !busy;
            btn_replace.IsEnabled = !busy;
            if (status != null)
            {
                txt_status.Text = status;
            }
        }

        private void SetStatus(string status) => txt_status.Text = status;
    }
}

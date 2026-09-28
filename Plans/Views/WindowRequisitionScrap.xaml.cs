using NLog;
using ORT一键报告.Models;
using ORT一键报告.Services;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Media;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// WindowRequisitionScrap.xaml 的交互逻辑：领退表右键菜单「报废登记」。
    /// 只读展示该条领退记录的机种/单据/序列号/工令/线别（可选中复制）；
    /// 报废必须提供序列号清单（直接输入或上传文件），点「报废」时与领用记录的清单
    /// 完全核对（多/少/重复都算异常并阻止）；通过后由调用方把报废字段写回该记录
    /// （走暂存 → 点「提交保存」时统一入库并写变更日志）或提交审核。窗口本身不写数据库。
    /// </summary>
    public partial class WindowRequisitionScrap : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;
        private readonly Requisition _req;
        private readonly bool _needsReview;

        /// <summary>打开窗口时预提取的领用序列号清单（核对基准）</summary>
        private List<string> _requisitionSns = [];

        /// <summary>文件模式选择的源文件路径</summary>
        private string _uploadedSnFile;

        /// <summary>报废单据号（DialogResult 为 true 时有效；未填为空）</summary>
        public string ScrapNo { get; private set; }

        /// <summary>报废数量（DialogResult 为 true 时有效）</summary>
        public string ScrapQty { get; private set; }

        /// <summary>报废日期（DialogResult 为 true 时有效）</summary>
        public DateTime ScrapDate { get; private set; }

        /// <summary>报废序列号清单文本（文本模式；文件模式为 null）</summary>
        public string ScrapSnText { get; private set; }

        /// <summary>报废序列号文件（文件模式，OleDir 相对文件名；文本模式为 null）</summary>
        public string ScrapSnFilePath { get; private set; }

        /// <summary>报废序列号个数（去重后）</summary>
        public int ScrapSnCount { get; private set; }

        /// <param name="db">数据库服务（解析领用附件路径 / 保存报废清单文件）</param>
        /// <param name="req">目标领退记录</param>
        /// <param name="needsReview">普通用户（提交审核）时提示文字不同</param>
        /// <param name="snDisplayText">序列号展示文本（附件形式时传文件名/路径）</param>
        public WindowRequisitionScrap(DatabaseService db, Requisition req, bool needsReview, string snDisplayText)
        {
            InitializeComponent();
            _db = db;
            _req = req;
            _needsReview = needsReview;

            txt_model.Text = req?.ModelName ?? "";
            txt_reqNo.Text = req?.RequisitionNo ?? "";
            txt_sn.Text = snDisplayText ?? "";
            txt_workOrder.Text = req?.WorkOrder ?? "";
            txt_line.Text = req?.LineNo ?? "";
            // 已有报废信息时沿用，方便修正；没有则日期默认当前日期
            txt_scrapNo.Text = req?.ScrapNo ?? "";
            txt_scrapQty.Text = req?.ScrapQty ?? "";
            dp_scrapDate.SelectedDate = req?.ScrapDate ?? DateTime.Today;
            txt_hint.Text = LanguageService.Get(needsReview ? "ReqScrap_HintReview" : "ReqScrap_Hint");

            LoadRequisitionSns();
        }

        /* ###############################  领用序列号清单  ################################ */

        /// <summary>
        /// 预提取领用记录的序列号清单并显示个数与来源；无法取得时显示红色提示（提交时同样阻止）
        /// </summary>
        private void LoadRequisitionSns()
        {
            try
            {
                string filePath = _db.ResolveAttachmentPath(_req?.SnFilePath);
                _requisitionSns = SnVerification.ExtractFromRequisition(_req?.SN, filePath);
            }
            catch (Exception ex)
            {
                _logger.Warn($"提取领用记录序列号失败: {ex.Message}");
                _requisitionSns = [];
            }

            string source = !string.IsNullOrWhiteSpace(_req?.SnFilePath)
                ? string.Format(LanguageService.Get("ReqScrap_SnFromFileFormat"), Path.GetFileName(_req.SnFilePath))
                : LanguageService.Get("ReqScrap_SnFromText");
            if (_requisitionSns.Count > 0)
            {
                txt_reqSnInfo.Text = string.Format(LanguageService.Get("ReqScrap_SnReqInfoFormat"),
                    _requisitionSns.Count, source);
                txt_reqSnInfo.Foreground = FindBrush("TextSecondaryBrush");
            }
            else
            {
                txt_reqSnInfo.Text = LanguageService.Get("Msg_ScrapSnReqUnavailable");
                txt_reqSnInfo.Foreground = FindBrush("StatusPendingBrush");
            }
        }

        private Brush FindBrush(string key)
            => TryFindResource(key) as Brush ?? SystemColors.ControlTextBrush;

        /* ###############################  输入模式  ################################ */

        private void SnMode_Changed(object sender, RoutedEventArgs e)
        {
            if (txt_scrapSn == null || btn_snFile == null)
            {
                return;
            }
            bool isInput = rb_snInput.IsChecked == true;
            txt_scrapSn.Visibility = isInput ? Visibility.Visible : Visibility.Collapsed;
            btn_snFile.Visibility = isInput ? Visibility.Collapsed : Visibility.Visible;
            txt_snFileName.Visibility = isInput ? Visibility.Collapsed : Visibility.Visible;
        }

        private void Btn_SnFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dialog = new()
            {
                Title = LanguageService.Get("Title_SelectSNFile"),
                Filter = "Excel文件|*.xls;*.xlsx;*.xlsm|文本文件|*.txt;*.csv|所有文件|*.*"
            };
            if (dialog.ShowDialog() == true)
            {
                _uploadedSnFile = dialog.FileName;
                txt_snFileName.Text = _uploadedSnFile;
            }
        }

        /* ###############################  报废确认  ################################ */

        private void Btn_Scrap_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_scrapQty.Text))
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillScrapQty"), LanguageService.Get("Cap_Info"));
                txt_scrapQty.Focus();
                return;
            }
            if (dp_scrapDate.SelectedDate is not DateTime date)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillScrapDate"), LanguageService.Get("Cap_Info"));
                return;
            }

            // 1. 取报废序列号清单
            bool fileMode = rb_snFile.IsChecked == true;
            List<string> scrapSns;
            if (fileMode)
            {
                if (string.IsNullOrWhiteSpace(_uploadedSnFile) || !File.Exists(_uploadedSnFile))
                {
                    _ = MessageBox.Show(LocalizationHelper.Get("Msg_SelectScrapSnFile"), LanguageService.Get("Cap_Info"));
                    return;
                }
                try
                {
                    scrapSns = SnVerification.ExtractFromFile(_uploadedSnFile);
                }
                catch (Exception ex)
                {
                    _ = MessageBox.Show(string.Format(LocalizationHelper.Get("Msg_ScrapSnExtractFailedFormat"), ex.Message),
                        LanguageService.Get("Cap_Info"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }
            else
            {
                scrapSns = SnVerification.ParseText(txt_scrapSn.Text);
                if (scrapSns.Count == 0)
                {
                    _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillScrapSn"), LanguageService.Get("Cap_Info"));
                    txt_scrapSn.Focus();
                    return;
                }
            }

            // 2. 与领用清单核对：异常则报告并阻止（窗口保持打开）
            SnVerification.SnCompareResult result = SnVerification.Compare(_requisitionSns, scrapSns);
            if (!result.Ok)
            {
                _ = MessageBox.Show(BuildAnomalyText(result), LanguageService.Get("Cap_ScrapSnCheckTitle"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // 3. 通过：文件模式复制文件到附件目录，文本模式保存文本
            if (fileMode)
            {
                string saved = SaveScrapSnFile(_uploadedSnFile);
                if (saved == null)
                {
                    return;
                }
                ScrapSnFilePath = saved;
                ScrapSnText = null;
            }
            else
            {
                ScrapSnText = string.Join(Environment.NewLine, scrapSns);
                ScrapSnFilePath = null;
            }
            ScrapSnCount = result.ScrapCount;
            ScrapNo = string.IsNullOrWhiteSpace(txt_scrapNo.Text) ? null : txt_scrapNo.Text.Trim();
            ScrapQty = txt_scrapQty.Text.Trim();
            ScrapDate = date;
            DialogResult = true;
        }

        private void Btn_Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;

        /* ###############################  核对结果文本  ################################ */

        /// <summary>
        /// 组装核对未通过的明细文本（分「不在领用清单 / 漏报废 / 清单内重复」三节，每节最多列 15 条）
        /// </summary>
        private static string BuildAnomalyText(SnVerification.SnCompareResult result)
        {
            if (result.RequisitionListEmpty)
            {
                return LocalizationHelper.Get("Msg_ScrapSnReqUnavailable");
            }
            if (result.ScrapListEmpty)
            {
                return LocalizationHelper.Get("Msg_FillScrapSn");
            }
            StringBuilder sb = new();
            sb.AppendLine(LocalizationHelper.Get("Msg_ScrapSnCheckHeader"));
            AppendSection(sb, "Msg_ScrapSnCheckNotIn", result.NotInRequisition, null);
            AppendSection(sb, "Msg_ScrapSnCheckMissing", result.MissingInScrap, null);
            AppendSection(sb, "Msg_ScrapSnCheckDuplicate", null,
                result.DuplicatesInScrap.Select(d => string.Format(LocalizationHelper.Get("Msg_ScrapSnDuplicateItemFormat"), d.Sn, d.Count)).ToList());
            return sb.ToString().TrimEnd();
        }

        private static void AppendSection(StringBuilder sb, string countFormatKey, List<string> items, List<string> rendered)
        {
            const int maxShown = 15;
            List<string> list = rendered ?? items ?? [];
            if (list.Count == 0)
            {
                return;
            }
            string shown = string.Join("、", list.Take(maxShown));
            if (list.Count > maxShown)
            {
                shown += Environment.NewLine + string.Format(LocalizationHelper.Get("Msg_ScrapSnCheckMoreFormat"), list.Count);
            }
            sb.AppendLine(string.Format(LocalizationHelper.Get(countFormatKey), list.Count) + shown);
        }

        /* ###############################  保存报废清单文件  ################################ */

        /// <summary>
        /// 把上传的报废序列号文件复制到附件目录（OleDir），命名与 SN 文件上传风格一致；
        /// 失败弹提示并返回 null
        /// </summary>
        private string SaveScrapSnFile(string sourcePath)
        {
            try
            {
                string key = _req?.RequisitionNo ?? "无单据号";
                string model = _req?.ModelName ?? "无机种名";
                string name = $"{DateTime.Now:MMdd}_報廢_{Clean(key)}_{Clean(model)}_{Clean(Path.GetFileName(sourcePath))}";
                string fullPath = Path.Combine(_db.OleDir, name);
                if (File.Exists(fullPath))
                {
                    name = $"{DateTime.Now:MMddHHmmss}_報廢_{Clean(key)}_{Clean(model)}_{Clean(Path.GetFileName(sourcePath))}";
                    fullPath = Path.Combine(_db.OleDir, name);
                }
                File.Copy(sourcePath, fullPath, true);
                return name;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "保存报废序列号文件失败");
                _ = MessageBox.Show($"保存报废序列号文件失败:\n{ex.Message}", LanguageService.Get("Cap_Error"));
                return null;
            }
        }

        private static string Clean(string name)
        {
            string cleaned = Regex.Replace(name ?? "", $"[{Regex.Escape(new string(Path.GetInvalidFileNameChars()))}]", "_").Trim();
            return cleaned == "" ? "_" : cleaned;
        }
    }
}

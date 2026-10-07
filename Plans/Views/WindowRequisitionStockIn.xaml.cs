using ORT一键报告.Models;
using ORT一键报告.Services;
using ORT一键报告.Utils;
using System;
using System.Windows;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// WindowRequisitionStockIn.xaml 的交互逻辑：领退表右键菜单「入库」。
    /// 只读展示该条领退记录的机种名称/领用单据/序列号/工令/线别（可选中复制），
    /// 用户填写入库单据/数量/日期后点「入库」，由调用方把这三个值写回该记录（走暂存 →
    /// 点「提交保存」时统一入库并写变更日志）。窗口本身不写数据库。
    /// 「下一个待入库」＝同样写入这三个值，并让调用方接着打开下一条待入库记录的登记窗口。
    /// </summary>
    public partial class WindowRequisitionStockIn : Window
    {
        /// <summary>入库单据（DialogResult 为 true 时有效）</summary>
        public string StockInNo { get; private set; }

        /// <summary>入库数量（DialogResult 为 true 时有效）</summary>
        public string StockInQty { get; private set; }

        /// <summary>入库日期（DialogResult 为 true 时有效）</summary>
        public DateTime StockInDate { get; private set; }

        /// <summary>是否点了「下一个待入库」：调用方据此接着打开下一条待入库记录（DialogResult 为 true 时有效）</summary>
        public bool NextRequested { get; private set; }

        /// <param name="req">目标领退记录</param>
        /// <param name="snText">序列号展示文本（附件形式时传文件名/路径）</param>
        public WindowRequisitionStockIn(Requisition req, string snText)
        {
            InitializeComponent();
            txt_model.Text = req?.ModelName ?? "";
            txt_reqNo.Text = req?.RequisitionNo ?? "";
            txt_sn.Text = snText ?? "";
            txt_workOrder.Text = req?.WorkOrder ?? "";
            txt_line.Text = req?.LineNo ?? "";
            // 已有入库信息时沿用，方便修正；没有则日期默认当前日期
            txt_stockInNo.Text = req?.StockInNo ?? "";
            txt_stockInQty.Text = req?.StockInQty ?? "";
            dp_stockInDate.SelectedDate = req?.StockInDate ?? DateTime.Today;
            // 上面几个只读值：右键单击即复制该值（入库单据/数量/日期是输入框，保持系统默认行为）
            RightClickCopy.AttachReadOnlyText(this);
        }

        private void Btn_StockIn_Click(object sender, RoutedEventArgs e)
        {
            if (ValidateInputs())
            {
                DialogResult = true;
            }
        }

        /// <summary>「下一个待入库」：这条的入库信息照样交出去，并让调用方接着打开下一条待入库记录</summary>
        private void Btn_NextPending_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateInputs())
            {
                return;
            }
            NextRequested = true;
            DialogResult = true;
        }

        /// <summary>校验入库单据/数量/日期并把值取到属性上；校验不通过时提示并聚焦，返回 false</summary>
        private bool ValidateInputs()
        {
            if (string.IsNullOrWhiteSpace(txt_stockInNo.Text))
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillStockInNo"), LanguageService.Get("Cap_Info"));
                txt_stockInNo.Focus();
                return false;
            }
            if (string.IsNullOrWhiteSpace(txt_stockInQty.Text))
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillStockInQty"), LanguageService.Get("Cap_Info"));
                txt_stockInQty.Focus();
                return false;
            }
            if (dp_stockInDate.SelectedDate is not DateTime date)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillStockInDate"), LanguageService.Get("Cap_Info"));
                return false;
            }
            StockInNo = txt_stockInNo.Text.Trim();
            StockInQty = txt_stockInQty.Text.Trim();
            StockInDate = date;
            return true;
        }

        private void Btn_Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;
    }
}

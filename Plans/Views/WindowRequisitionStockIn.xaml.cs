using ORT一键报告.Models;
using ORT一键报告.Services;
using System;
using System.Windows;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// WindowRequisitionStockIn.xaml 的交互逻辑：领退表右键菜单「入库」。
    /// 只读展示该条领退记录的机种名称/领用单据/序列号/工令/线别（可选中复制），
    /// 用户填写入库单据/数量/日期后点「入库」，由调用方把这三个值写回该记录（走暂存 →
    /// 点「提交保存」时统一入库并写变更日志）。窗口本身不写数据库。
    /// </summary>
    public partial class WindowRequisitionStockIn : Window
    {
        /// <summary>入库单据（DialogResult 为 true 时有效）</summary>
        public string StockInNo { get; private set; }

        /// <summary>入库数量（DialogResult 为 true 时有效）</summary>
        public string StockInQty { get; private set; }

        /// <summary>入库日期（DialogResult 为 true 时有效）</summary>
        public DateTime StockInDate { get; private set; }

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
        }

        private void Btn_StockIn_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_stockInNo.Text))
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillStockInNo"), LanguageService.Get("Cap_Info"));
                txt_stockInNo.Focus();
                return;
            }
            if (string.IsNullOrWhiteSpace(txt_stockInQty.Text))
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillStockInQty"), LanguageService.Get("Cap_Info"));
                txt_stockInQty.Focus();
                return;
            }
            if (dp_stockInDate.SelectedDate is not DateTime date)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillStockInDate"), LanguageService.Get("Cap_Info"));
                return;
            }
            StockInNo = txt_stockInNo.Text.Trim();
            StockInQty = txt_stockInQty.Text.Trim();
            StockInDate = date;
            DialogResult = true;
        }

        private void Btn_Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;
    }
}

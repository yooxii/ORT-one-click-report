using ORT一键报告.Models;
using ORT一键报告.Services;
using System;
using System.Windows;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// WindowRequisitionReturn.xaml 的交互逻辑：领退表右键菜单「回线」。
    /// 只读展示该条领退记录的机种名称/回线RT工令/回线数/线别（可选中复制），
    /// 日期默认当前日期；点「回线」后由调用方把所选日期写回该记录的回线日期（走暂存 →
    /// 点「提交保存」时统一入库并写变更日志）。窗口本身不写数据库。
    /// </summary>
    public partial class WindowRequisitionReturn : Window
    {
        /// <summary>确认回线时选择的日期（DialogResult 为 true 时有效）</summary>
        public DateTime ReturnDate { get; private set; }

        public WindowRequisitionReturn(Requisition req)
        {
            InitializeComponent();
            txt_model.Text = req?.ModelName ?? "";
            txt_returnRt.Text = req?.ReturnRtOrder ?? "";
            txt_returnQty.Text = req?.ReturnQty ?? "";
            txt_line.Text = req?.LineNo ?? "";
            // 日期默认当前日期；该记录已有回线日期时沿用，方便修正
            dp_returnDate.SelectedDate = req?.ReturnDate ?? DateTime.Today;
        }

        private void Btn_Return_Click(object sender, RoutedEventArgs e)
        {
            if (dp_returnDate.SelectedDate is not DateTime date)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillReturnDate"), LanguageService.Get("Cap_Info"));
                return;
            }
            ReturnDate = date;
            DialogResult = true;
        }

        private void Btn_Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;
    }
}

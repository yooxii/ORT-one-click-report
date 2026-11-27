using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Models;
using ORT一键报告.Services;
using ORT一键报告.Services.Sim44;
using ORT一键报告.Utils;
using System;
using System.Windows;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// WindowRequisitionReturn.xaml 的交互逻辑：领退表右键菜单「回线」。
    /// 只读展示该条领退记录的机种名称/回线RT工令/线别（可选中复制），
    /// 回线数量可编辑（默认为领用数量）；日期默认当前日期；
    /// 点「回线」后由调用方把所选日期与数量写回该记录（走暂存 →
    /// 点「提交保存」时统一入库并写变更日志）。窗口本身不写数据库。
    /// 另外提供「登记 MES 工令」入口：把同一份信息写进 MES 的 SIM44（原本要在 C_SIM19 里手敲）。
    /// </summary>
    public partial class WindowRequisitionReturn : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly Requisition _req;

        /// <summary>确认回线时选择的日期（DialogResult 为 true 时有效）</summary>
        public DateTime ReturnDate { get; private set; }

        /// <summary>确认回线时的回线数量（默认为领用数量，可编辑；DialogResult 为 true 时有效）</summary>
        public string ReturnQty { get; private set; }

        /// <summary>领用数量：回线数量留空时用它兜底</summary>
        private readonly string _outQty;

        public WindowRequisitionReturn(Requisition req)
        {
            InitializeComponent();
            _req = req;
            _outQty = req?.OutQty ?? "";
            txt_model.Text = req?.ModelName ?? "";
            txt_returnRt.Text = req?.ReturnRtOrder ?? "";
            // 回线数量：已经登记过的沿用，还没登记时默认领用数量
            txt_returnQty.Text = string.IsNullOrWhiteSpace(req?.ReturnQty) ? _outQty : req.ReturnQty;
            txt_line.Text = req?.LineNo ?? "";
            // 日期默认当前日期；该记录已有回线日期时沿用，方便修正
            dp_returnDate.SelectedDate = req?.ReturnDate ?? DateTime.Today;
            // 上面几个只读值：右键单击即复制该值（回线数量可编辑，保持系统默认右键菜单）
            RightClickCopy.AttachReadOnlyText(this);
            // 没有回线RT工令就没法登记 MES 工令，直接把入口置灰并在提示里说明
            bool hasRtOrder = !string.IsNullOrWhiteSpace(req?.ReturnRtOrder);
            btn_mes.IsEnabled = hasRtOrder;
            if (!hasRtOrder)
            {
                txt_mesState.Text = LanguageService.Get("ReqReturn_MesNeedRtOrder");
            }
        }

        private void Btn_Return_Click(object sender, RoutedEventArgs e)
        {
            if (dp_returnDate.SelectedDate is not DateTime date)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillReturnDate"), LanguageService.Get("Cap_Info"));
                return;
            }
            ReturnDate = date;
            string qty = txt_returnQty.Text?.Trim();
            ReturnQty = string.IsNullOrEmpty(qty) ? _outQty : qty;
            DialogResult = true;
        }

        private void Btn_Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;

        /// <summary>
        /// 打开 MES 回线工令维护窗口：工令=回线RT工令、线别/日期/数量按当前填写值预填，
        /// 机种名称作为「查料号」的提示带过去。
        /// </summary>
        private void Btn_Mes_Click(object sender, RoutedEventArgs e)
        {
            Sim44Service service = App.ServiceProvider.GetRequiredService<Sim44Service>();
            DateTime date = dp_returnDate.SelectedDate ?? DateTime.Today;
            string qty = txt_returnQty.Text?.Trim();
            if (string.IsNullOrEmpty(qty))
            {
                qty = _outQty;
            }
            Sim44Row prefill = new()
            {
                MoNo = _req?.ReturnRtOrder ?? "",
                LineNo = _req?.LineNo ?? "",
                StartDate = date.ToString("yyyy/MM/dd"),
                MoQty = qty ?? "",
            };
            WindowSim44Return window = new(service, prefill, _req?.ModelName) { Owner = this };
            window.ShowDialog();
            if (window.AnyChangeSaved)
            {
                txt_mesState.Text = LanguageService.Get("ReqReturn_MesSaved");
                _logger.Info($"回线：{_req?.RequisitionNo} 已同步登记 MES 工令 {prefill.MoNo}");
            }
        }
    }
}

using ORT一键报告.Models;
using ORT一键报告.Services;
using ORT一键报告.Utils;
using System;
using System.Windows;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// WindowPlanUnitReturn.xaml 的交互逻辑：计划表右键菜单「单体归还」（其他部门申请测试流程）。
    /// 只读展示该计划的工作编号/机种名称/测试项目/样品数（可选中复制），
    /// 归还日期默认当前日期；点「单体归还」后由调用方把所选日期写回该计划的归还日期
    /// （走暂存 → 点「提交保存」时统一入库并写变更日志）。窗口本身不写数据库。
    /// </summary>
    public partial class WindowPlanUnitReturn : Window
    {
        /// <summary>确认归还时选择的日期（DialogResult 为 true 时有效）</summary>
        public DateTime UnitReturnDate { get; private set; }

        public WindowPlanUnitReturn(Plan plan)
        {
            InitializeComponent();
            txt_jobNo.Text = plan?.JobNo ?? "";
            txt_model.Text = plan?.ModelName ?? "";
            txt_testItem.Text = plan?.TestItem ?? "";
            txt_sampleSize.Text = plan?.SampleSize ?? "";
            // 日期默认当前日期；该计划已有归还日期时沿用，方便修正
            dp_returnDate.SelectedDate = plan?.UnitReturnDate ?? DateTime.Today;
            // 上面几个只读值：右键单击即复制该值
            RightClickCopy.AttachReadOnlyText(this);
        }

        private void Btn_Return_Click(object sender, RoutedEventArgs e)
        {
            if (dp_returnDate.SelectedDate is not DateTime date)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillUnitReturnDate"), LanguageService.Get("Cap_Info"));
                return;
            }
            UnitReturnDate = date;
            DialogResult = true;
        }

        private void Btn_Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;
    }
}

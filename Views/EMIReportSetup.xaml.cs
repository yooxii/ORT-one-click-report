using System.Windows;

namespace ORT一键报告
{
    /// <summary>
    /// EMIReportSetup.xaml 的交互逻辑
    /// </summary>
    public partial class EMIReportSetup : Window
    {
        public EMIReportSetup()
        {
            InitializeComponent();
            DataContext = EMIReportPage.emiVM;
            EMIReportPage.emiVM.LoadFromExcel();
        }
    }
}
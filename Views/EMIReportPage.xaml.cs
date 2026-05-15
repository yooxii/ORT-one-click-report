using System.Windows.Controls;

namespace ORT一键报告
{

    /// <summary>
    /// EMIReportPage.xaml 的交互逻辑
    /// </summary>
    public partial class EMIReportPage : UserControl
    {
        public static ViewModels.EMIViewModel emiVM;
        public EMIReportPage()
        {
            InitializeComponent();
            EMIService service = new();
            emiVM = new(service);
            DataContext = emiVM;
        }
    }
}

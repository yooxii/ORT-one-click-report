using ORT一键报告.ViewModels;
using System.Windows.Controls;

namespace ORT一键报告.Views
{
    /// <summary>
    /// UserControlReturnLineSingle.xaml 的交互逻辑
    /// </summary>
    public partial class UserControlReturnLineSingle : UserControl
    {
        public ReturnLineSingleViewModel RLSVM { get; set; }
        public UserControlReturnLineSingle()
        {
            InitializeComponent();
            DataContext = RLSVM;
        }
    }
}

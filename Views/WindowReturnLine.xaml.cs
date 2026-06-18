using ORT一键报告.ViewModels;
using System.Collections.Generic;
using System.Windows;

namespace ORT一键报告.Views
{
    /// <summary>
    /// WindowReturnLine.xaml 的交互逻辑
    /// </summary>
    public partial class WindowReturnLine : Window
    {
        public List<ReturnLineSingleViewModel> RLVM { get; set; }
        public WindowReturnLine()
        {
            InitializeComponent();
        }
    }
}

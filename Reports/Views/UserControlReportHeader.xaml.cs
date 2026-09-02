using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ORT一键报告.Reports.ViewModels;

namespace ORT一键报告.Reports.Views
{
    /// <summary>
    /// ReportHeader.xaml 的交互逻辑
    /// </summary>
    public partial class ReportHeaderWidget : UserControl
    {
        public int TestTime
        {
            get => (int)GetValue(TestTimeProperty);
            set => SetValue(TestTimeProperty, value);
        }

        public static readonly DependencyProperty TestTimeProperty =
            DependencyProperty.Register("TestTime", typeof(int), typeof(ReportHeaderWidget), new PropertyMetadata(1));


        public string TestedBy { set => SetValue(TestedByProperty, value); get => (string)GetValue(TestedByProperty); }
        public static readonly DependencyProperty TestStageProperty = DependencyProperty.Register(nameof(TestStage), typeof(string), typeof(ReportHeaderWidget), new PropertyMetadata(string.Empty));

        public string ProjectName { set => SetValue(ProjectNameProperty, value); get => (string)GetValue(ProjectNameProperty); }
        public static readonly DependencyProperty ProjectNameProperty = DependencyProperty.Register(nameof(ProjectName), typeof(string), typeof(ReportHeaderWidget), new PropertyMetadata(string.Empty));

        public string ApprovedBy { set => SetValue(ApprovedByProperty, value); get => (string)GetValue(ApprovedByProperty); }
        public static readonly DependencyProperty ApprovedByProperty = DependencyProperty.Register(nameof(ApprovedBy), typeof(string), typeof(ReportHeaderWidget), new PropertyMetadata(string.Empty));

        public string TestStage { set => SetValue(TestStageProperty, value); get => (string)GetValue(TestStageProperty); }
        public static readonly DependencyProperty TestedByProperty = DependencyProperty.Register(nameof(TestedBy), typeof(string), typeof(ReportHeaderWidget), new PropertyMetadata(string.Empty));

        public string TextTestDescription
        {
            get => (string)GetValue(TextTestDescriptionProperty);
            set => SetValue(TextTestDescriptionProperty, value);
        }
        public static readonly DependencyProperty TextTestDescriptionProperty =
            DependencyProperty.Register("TextTestDescription", typeof(string), typeof(ReportHeaderWidget), new PropertyMetadata(string.Empty));

        public ImageSource PicTestDescription
        {
            get => (ImageSource)GetValue(PicTestDescriptionProperty);
            set => SetValue(PicTestDescriptionProperty, value);
        }
        public static readonly DependencyProperty PicTestDescriptionProperty =
            DependencyProperty.Register("PicTestDescription", typeof(ImageSource), typeof(ReportHeaderWidget), new PropertyMetadata(null));

        public ReportHeaderWidget()
        {
            InitializeComponent();
        }
        private void Datepicker_start_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is DatePicker datepicker_start && datepicker_start.SelectedDate is DateTime startDate)
            {
                DateTime endDate = startDate.AddDays(TestTime);
                datepicker_end.SelectedDate = endDate;
                if (DataContext is ReportHeaderViewModel vm)
                {
                    vm.TestStart = startDate;
                    vm.TestEnd = endDate;
                }
            }
        }
    }
}

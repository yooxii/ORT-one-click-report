using ORT一键报告.ViewModels;
using System.Windows;
using System.Windows.Controls;

namespace ORT一键报告.Views
{
    /// <summary>
    /// SettingControl.xaml 的交互逻辑
    /// </summary>
    public partial class SettingControl : UserControl
    {
        public SettingItemViewModel SettingItemVM
        {
            get { return (SettingItemViewModel)GetValue(SettingItemVMProperty); }
            set { SetValue(SettingItemVMProperty, value); }
        }

        // Using a DependencyProperty as the backing store for SettingItemVM.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty SettingItemVMProperty =
            DependencyProperty.Register(nameof(SettingItemVM), typeof(SettingItemViewModel), typeof(SettingControl), new PropertyMetadata(null));



        public SettingControl()
        {
            InitializeComponent();
            Loaded += Control_Loading;
        }
        private void Control_Loading(object sender, RoutedEventArgs e)
        {
            if (SettingItemVM != null)
            {
                DataContext = SettingItemVM;
                if (SettingItemVM.Children.Count > 0)
                {
                    foreach (var vm in SettingItemVM.Children)
                    {
                        contentControl.Children.Add(new SettingControl() { SettingItemVM = vm });
                    }
                }
            }
            else
            {
                MessageBox.Show("SettingControl的DataContext为空", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }
}

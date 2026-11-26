using ORT一键报告.Services;
using System.Windows;

namespace ORT一键报告.Main.Views
{
    /// <summary>
    /// 看板平台配置窗口：把「ORT 看板取数」设置单独开出来，
    /// 入口在领退窗口的「工具」菜单，权限放到技术员及以上（不需要管理员）。
    /// 界面就是设置窗口里的同一份 <see cref="UserControlOrtSettings"/>，这里只多一对保存/关闭按钮。
    /// </summary>
    public partial class WindowOrtSettings : Window
    {
        public WindowOrtSettings()
        {
            InitializeComponent();
            Loaded += (s, e) => uc_ort.Load();
        }

        private void Btn_Save_Click(object sender, RoutedEventArgs e)
        {
            if (!uc_ort.Apply())
            {
                return;
            }
            ToastService.Show(LanguageService.Get("Ort_Settings_Saved"), ToastType.Info);
        }

        private void Btn_Close_Click(object sender, RoutedEventArgs e) => Close();
    }
}

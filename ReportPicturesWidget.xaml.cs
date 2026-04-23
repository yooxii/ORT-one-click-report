using Microsoft.Win32;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ORT一键报告
{
    /// <summary>
    /// ReportPicturesWidget.xaml 的交互逻辑
    /// </summary>
    public partial class ReportPicturesWidget : UserControl
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public ReportPicturesWidget()
        {
            InitializeComponent();
        }

        private void Label_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is Label label && label.Tag is string imageName)
            {
                FileDialog img = new OpenFileDialog
                {
                    Filter = "图片文件|*.png;*.jpg;*.jpeg;*.gif"
                };
                _ = img.ShowDialog();

                if (img.FileName is string img_path && img_path != "")
                {
                    if (FindName(imageName) is Image image)
                    {
                        image.Source = new BitmapImage(new Uri(img_path));
                    }
                }
                else
                {
                    _logger.Warn("未选择图片");
                    _ = MessageBox.Show("未选择图片");
                }
            }
        }

        private void Image_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (sender is Image image)
            {
                FileDialog imgDialog = new OpenFileDialog()
                {
                    Filter = "图片|*.png;*.jpg;*.jepg",
                };
                if (imgDialog.ShowDialog() == true)
                {
                    string imgName = imgDialog.FileName;
                    image.Source = new BitmapImage(new Uri(imgName));
                }
            }
            else
            {
                _logger.Warn("未选择图片");
                _ = MessageBox.Show("未选择图片");
            }
        }
    }
}

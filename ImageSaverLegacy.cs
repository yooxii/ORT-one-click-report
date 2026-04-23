using System;
using System.IO;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace ORT一键报告
{
    public static class ImageSaverLegacy
    {
        /// <summary>
        /// 将 WPF ImageSource 保存为文件
        /// </summary>
        public static void SaveImageSourceToFile(ImageSource imageSource, string filePath, string format)
        {
            if (imageSource == null)
            {
                throw new ArgumentNullException(nameof(imageSource));
            }

            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException(nameof(filePath));
            }

            // 1. 确保目录存在
            string directory = System.IO.Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory))
            {
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }
            }

            string formatLower = format.ToLower();


            // 2. 创建编码器
            BitmapEncoder encoder;
            switch (formatLower)
            {
                case "jpg":
                case "jpeg":
                    encoder = new JpegBitmapEncoder();
                    break;
                case "bmp":
                    encoder = new BmpBitmapEncoder();
                    break;
                case "gif":
                    encoder = new GifBitmapEncoder();
                    break;
                case "tiff":
                    encoder = new TiffBitmapEncoder();
                    break;
                case "png":
                default:
                    encoder = new PngBitmapEncoder();
                    break;
            }

            // 3. 设置 JPEG 质量 (可选)
            if (encoder is JpegBitmapEncoder encoder1)
            {
                encoder1.QualityLevel = 90;
            }

            // 4. 添加帧
            // BitmapFrame.Create 会处理 ImageSource 到帧的转换
            encoder.Frames.Add(BitmapFrame.Create((BitmapSource)imageSource));

            // 5. 写入文件
            using (FileStream stream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                encoder.Save(stream);
            }
        }
    }
}

using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Text;

namespace ORT一键报告
{
    public static class ImageUtils
    {
        /// <summary>
        /// 生成紧凑布局的 EMF 图片
        /// </summary>
        /// <param name="filePath">保存的 EMF 文件路径</param>
        /// <param name="icon">已加载的图标对象</param>
        /// <param name="text">要绘制的字符串</param>
        /// <param name="margin">边距（默认2px，防止边缘被裁切）</param>
        public static void GenerateCenteredEmf(string filePath, Bitmap icon, string text, int margin = 2)
        {
            if (icon == null) throw new ArgumentNullException(nameof(icon));

            try
            {
                // 1. 定义字体
                Font font = new("Microsoft YaHei", 12, FontStyle.Bold);
                float spacing = 10f; // 图标与文字之间的间距

                // 2. 获取图标尺寸
                float iconW = icon.Width;
                float iconH = icon.Height;

                // 3. 创建临时 Bitmap 用于获取 HDC 和测量文字
                using (Bitmap tempBmp = new(1, 1))
                using (Graphics tempGs = Graphics.FromImage(tempBmp))
                {
                    tempGs.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

                    // 4. 动态测量字符串尺寸（向上取整，防止文字边缘被裁切）
                    SizeF textSize = tempGs.MeasureString(text, font);
                    float textW = (float)Math.Ceiling(textSize.Width);
                    float textH = (float)Math.Ceiling(textSize.Height);

                    // 5. 计算垂直布局的整体内容尺寸
                    float contentWidth = Math.Max(iconW, textW);
                    float contentHeight = iconH + spacing + textH;

                    // 6. 计算自适应画布尺寸（内容尺寸 + 极小的安全边距）
                    int canvasWidth = (int)Math.Ceiling(contentWidth + margin * 2);
                    int canvasHeight = (int)Math.Ceiling(contentHeight + margin * 2);

                    // 7. 创建目标尺寸的 Metafile
                    using Metafile mf = new(filePath, tempGs.GetHdc());
                    using Graphics g = Graphics.FromImage(mf);
                    // 8. 设置高质量渲染与白色背景
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                    g.Clear(Color.White);

                    // 9. 内容起点即为 margin
                    float startX = margin;
                    float startY = margin;

                    // 10. 绘制图标（水平居中于整体内容宽度）
                    float iconX = startX + (contentWidth - iconW) / 2f;
                    g.DrawImage(icon, iconX, startY);

                    // 11. 绘制脚注文字（水平居中于整体内容宽度，垂直位于图标下方）
                    float textX = startX + (contentWidth - textW) / 2f;
                    float textY = startY + iconH + spacing;
                    g.DrawString(text, font, Brushes.Black, textX, textY);

                    // 10. 保存并释放资源
                    g.Save();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"生成 EMF 失败: {ex.Message}");
            }
        }
    }
}

using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;

namespace ORT一键报告.Converters
{
    /// <summary>
    /// 布尔 → 文字样式转换器：审核详情里「与库中现值不同」的字段用高亮样式（底色 + 强调色文字），
    /// 其余用普通样式。样式在这里直接构造，不依赖资源字典查找，避免窗口/应用级资源查找失败。
    /// </summary>
    public class ChangedStyleConverter : IValueConverter
    {
        /// <summary>高亮底色（对应主题 StatusWarnBgBrush）</summary>
        private static readonly Brush ChangedBackground = ChangedBrush.Background;

        /// <summary>高亮文字色（对应主题 StatusWarnBrush）</summary>
        private static readonly Brush ChangedForeground = ChangedBrush.Foreground;

        private static readonly Style NormalStyle = BuildStyle(false);

        private static readonly Style HighlightStyle = BuildStyle(true);

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => value is bool changed && changed ? HighlightStyle : NormalStyle;

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();

        /// <summary>
        /// 构造 TextBlock 样式：普通样式只统一行距，高亮样式再加底色与强调色
        /// </summary>
        private static Style BuildStyle(bool highlight)
        {
            Style style = new(typeof(TextBlock));
            style.Setters.Add(new Setter(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Top));
            style.Setters.Add(new Setter(TextBlock.TextWrappingProperty, TextWrapping.Wrap));
            style.Setters.Add(new Setter(FrameworkElement.MarginProperty,
                highlight ? new Thickness(0, 2, 0, 2) : new Thickness(0, 3, 0, 3)));
            if (highlight)
            {
                style.Setters.Add(new Setter(Control.BackgroundProperty, ChangedBackground));
                style.Setters.Add(new Setter(Control.ForegroundProperty, ChangedForeground));
                style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(4, 1, 4, 1)));
            }
            style.Seal();
            return style;
        }
    }

    /// <summary>
    /// 布尔 → 画刷转换器：变动字段的整行底色，未变动返回透明（避免整表都是色块）。
    /// </summary>
    public class ChangedBackgroundConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => value is bool changed && changed ? ChangedBrush.Background : Brushes.Transparent;

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }

    /// <summary>
    /// 布尔 → Visibility 转换器：true 显示、false 折叠。
    /// </summary>
    public class BoolToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => value is bool flag && flag ? Visibility.Visible : Visibility.Collapsed;

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }

    /// <summary>高亮配色的唯一定义处，供上面几个转换器共用，避免各写一份而走样</summary>
    internal static class ChangedBrush
    {
        /// <summary>高亮底色（对应主题 StatusWarnBgBrush）</summary>
        internal static readonly Brush Background = new SolidColorBrush(Color.FromRgb(0xFF, 0xF3, 0xE0));

        /// <summary>高亮文字色（对应主题 StatusWarnBrush）</summary>
        internal static readonly Brush Foreground = new SolidColorBrush(Color.FromRgb(0xE6, 0x51, 0x00));
    }
}
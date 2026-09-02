using System;
using System.Globalization;
using System.Windows.Data;

namespace ORT一键报告.Converters
{
    /// <summary>
    /// 布尔值取反转换器：用于 FAIL 单选按钮与 TestPass 的双向绑定。
    /// </summary>
    public class InverseBoolConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => !(value is bool b && b);

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => value is bool b && !b;
    }
}

using System;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows.Data;

namespace ORT一键报告.Converters
{
    /// <summary>
    /// 多行文本压成单行显示（表格单元格用）：换行与连续空白折成一个空格，
    /// 避免 S/N 清单、备注这类多行内容把行高撑开；空值显示为空串。
    /// </summary>
    public class SingleLineTextConverter : IValueConverter
    {
        private static readonly Regex Whitespace = new(@"\s+", RegexOptions.Compiled);

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null)
            {
                return "";
            }
            string text = value as string ?? value.ToString();
            return Whitespace.Replace(text.Replace('\r', ' ').Replace('\n', ' '), " ").Trim();
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => Binding.DoNothing;
    }
}

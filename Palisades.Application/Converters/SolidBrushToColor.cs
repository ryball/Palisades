using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace Palisades.Converters
{
    internal class SolidBrushToColor : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value switch
            {
                SolidColorBrush colorBrush => colorBrush.Color,
                Color color => color,
                _ => Colors.White
            };
        }

        public object? ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Color color)
            {
                return new SolidColorBrush(color);
            }

            if (value is SolidColorBrush brush)
            {
                return brush;
            }

            return new SolidColorBrush(Colors.White);
        }
    }
}

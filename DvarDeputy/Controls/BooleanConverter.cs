using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace Mooseware.DvarDeputy.Controls;

public class BooleanConverter<T>(T trueValue, T falseValue) : IValueConverter
{
    public T True { get; set; } = trueValue;
    public T False { get; set; } = falseValue;

    public virtual object? Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is bool boolean && boolean ? True : False;
    }

    public virtual object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is T t && EqualityComparer<T>.Default.Equals(t, True);
    }
}

public sealed class BooleanToVisibilityConverter : BooleanConverter<Visibility>
{
    public BooleanToVisibilityConverter() :
    base(Visibility.Visible, Visibility.Collapsed)
    { }
}

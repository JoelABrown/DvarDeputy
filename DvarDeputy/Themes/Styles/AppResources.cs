using System.Windows;
using System.Windows.Media;

namespace Mooseware.DvarDeputy.Themes.Styles;

/// <summary>
/// Application specific styling resources
/// </summary>
internal static class AppResources
{
    internal enum StaticResource
    {
        BrightForegroundBrush,
        CountdownNormalBackgroundBrush,
        CountdownNormalBackgroundBorderBrush,
        CountdownWarningBrush,
        CountdownWarningBackgroundBrush,
        CountdownWarningBackgroundBorderBrush,
        CuedBackgroundBrush,
        CuedBackgroundBorderBrush,
        CuedContrastBrush,
        CuedMainBrush,
        PlayingBackgroundBrush,
        PlayingBackgroundBorderBrush,
        PlayingMainBrush,
        PlayingContrastBrush,
        DisabledMainBrush,
        DisabledContrastBrush,
        NoClipMainBrush,
        NoClipBackgroundBrush,
        NoClipBackgroundBorderBrush
    }

    internal static Brush DefinedColour(StaticResource colour)
    {
        // NOTE: This relies on the Enum name being the same as the StaticResource x:Key
        return (Brush)Application.Current.Resources[colour.ToString()];
    }
}

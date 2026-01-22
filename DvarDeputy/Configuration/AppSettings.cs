namespace Mooseware.DvarDeputy.Configuration;

/// <summary>
/// Application configuration settings (read-only at application start) for configuring the application
/// The settings are loaded from appsettings.json in a section called "ApplicationSettings".
/// </summary>
public class AppSettings
{
    /// <summary>
    /// The number of ticks in the TimerTickInterval for the timer that controls scrolling in the viewer window.
    /// </summary>
    public int TimerFrequency { get; set; } = 0;
    /// <summary>
    /// The lower limit for scrolling velocity (in pixels per timer tick)
    /// </summary>
    public double LowerScrollLimit { get; set; } = 0.0;
    /// <summary>
    /// The upper limit for scrolling velocity (in pixels per timer tick)
    /// </summary>
    public double UpperScrollLimit { get; set; } = double.MaxValue;
    /// <summary>
    /// The scrolling velocity set when the reset to default action is triggered
    /// </summary>
    public double DefaultScrollVelocity { get; set; } = 0.0;
    /// <summary>
    /// The lower limit for the font size in ems
    /// </summary>
    public int LowerFontSizeLimit { get; set; } = 0;
    /// <summary>
    /// The upper limit for the font size in ems
    /// </summary>
    public int UpperFontSizeLimit { get; set; } = int.MaxValue;
    /// <summary>
    /// The font size in ems which is set when the reset to default action is triggered
    /// </summary>
    public int DefaultFontSize { get; set; } = 0;
    /// <summary>
    /// The upper limit on the left and right margins in the viewer window (in pixels)
    /// </summary>
    public int UpperHorizontalMarginLimit { get; set; } = int.MaxValue;
    /// <summary>
    /// The left and right margins in the viewer window to set when the reset to default action is triggered
    /// </summary>
    public int DefaultHorizontalMargin { get; set; } = 0;
    /// <summary>
    /// The upper limit on the space between lines in the viewer window set as a multiple of line height
    /// </summary>
    public int UpperSpacingLimit { get; set; } = int.MaxValue; 
    /// <summary>
    /// The space between lines to set, in terms of multiples of line height, when the reset to default action is triggered
    /// </summary>
    public int DefaultLineSpacing { get; set;} = 0;
    /// <summary>
    /// The base of the URL for the HTTP Web API
    /// </summary>
    public string WebApiUrlRoot { get; set; } = "0.0.0.0";
    /// <summary>
    /// The port number of the URL for the HTTP Web API
    /// </summary>
    public string WebApiUrlPort { get; set; } = "80";
    /// <summary>
    /// The short name of the ATEM output to which the viewer is routed
    /// </summary>
    public string ViewerAtemOutput { get; set; } = "PGM";
}

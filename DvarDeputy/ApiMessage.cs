namespace Mooseware.DvarDeputy;

/// <summary>
/// Web API Message Container
/// </summary>
public enum ApiMessageVerb
{
    /// <summary>
    /// No API message defined (yet)
    /// </summary>
    None = 0,
    /// <summary>
    /// Viewer controls (e.g. show|hide)
    /// </summary>
    Viewer,
    /// <summary>
    /// Paging on the Viewer (e.g. previous|next|first)
    /// </summary>
    Page,
    /// <summary>
    /// Scrolling contol (e.g. forward|backward|stop)
    /// </summary>
    Scroll,
    /// <summary>
    /// Font sizing (e.g. increase|decrease|reset)
    /// Values other than 'reset' require a Scale parameter
    /// </summary>
    Font,
    /// <summary>
    /// Scrolling speed control (slower|faster|reset)
    /// Values other than 'reset' require a Scale parameter
    /// </summary>
    ScrollSpeed,
    /// <summary>
    /// Line spacing control (e.g. increase|decrease|reset)
    /// Values other than 'reset' require a Scale parameter
    /// </summary>
    Spacing,
    /// <summary>
    /// Side margins control (e.g. increase|decrease|reset)
    /// Values other than 'reset' require a Scale parameter
    /// </summary>
    Margin,
    /// <summary>
    /// Visual theme control (e.g. light|dark)
    /// </summary>
    Theme,
    /// <summary>
    /// Progress bug control (e.g. none|side|bottom)
    /// </summary>
    Bug
}

/// <summary>
/// An API message from a remote controlling app
/// </summary>
public class ApiMessage
{
    /// <summary>
    /// The action type of the message
    /// </summary>
    public ApiMessageVerb Verb { get; set; } = ApiMessageVerb.None;
    /// <summary>
    /// Particular details of what has been requested within the context of the type of message
    /// </summary>
    public string Parameters { get; set; } = string.Empty;
    /// <summary>
    /// A scalar amount (numeric type is context dependent) indicating "how much" to Verb/Parameter
    /// </summary>
    public double Scalar { get; set; } = 0.0;

    // String constants that define the acceptable query parameters for the various verbs
    // ----------------------------------------------------------------------------------

    public const string ViewerShow = "show";
    public const string ViewerHide = "hide";
    public const string PagePrevious = "previous";
    public const string PageNext = "next";
    public const string PageFirst = "first";
    public const string ScrollForward = "forward";
    public const string ScrollBackward = "backward";
    public const string ScrollStop = "stop";
    public const string FontIncrease = "increase";
    public const string FontDecrease = "decrease";
    public const string FontReset = "reset";
    public const string ScrollSpeedIncrease = "increase";
    public const string ScrollSpeedDecrease = "decrease";
    public const string ScrollSpeedReset = "reset";
    public const string SpacingIncrease = "increase";
    public const string SpacingDecrease = "decrease";
    public const string SpacingReset = "reset";
    public const string MarginIncrease = "increase";
    public const string MarginDecrease = "decrease";
    public const string MarginReset = "reset";
    public const string ThemeLight = "light";
    public const string ThemeDark = "dark";
    public const string ThemeMatrix = "matrix";
    public const string BugNone = "none";
    public const string BugSide = "side";
    public const string BugBottom = "bottom";
}
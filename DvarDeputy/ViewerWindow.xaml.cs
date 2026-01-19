using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace Mooseware.DvarDeputy;

/// <summary>
/// Interaction logic for ViewerWindow.xaml
/// </summary>
public partial class ViewerWindow : Window
{
    /// <summary>
    /// The window mode of the viewer
    /// </summary>
    public enum ViewerMode
    {
        /// <summary>
        /// Restored to normal size (movable, sizable, has chrome)
        /// </summary>
        Normal,
        /// <summary>
        /// Maximized to full screen with chrome hidden
        /// </summary>
        Fullscreen,
        /// <summary>
        /// Hidden so it is out of the way when not needed
        /// </summary>
        Hidden
    }

    /// <summary>
    /// The visual theme for the viewer window and prompted text
    /// </summary>
    public enum VisualTheme
    {
        /// <summary>
        /// Black text on a white background
        /// </summary>
        Light,
        /// <summary>
        /// White text on a black background
        /// </summary>
        Dark,
        /// <summary>
        /// Bright green text on a dark grey background
        /// </summary>
        Matrix
    }

    /// <summary>
    /// The location for showing a progress bug on the prompted text viewer
    /// </summary>
    public enum ProgressBug
    {
        /// <summary>
        /// Do not show a progress bug
        /// </summary>
        None,
        /// <summary>
        /// Show vertically on the right hand side of the viewer window
        /// </summary>
        RightSide,
        /// <summary>
        /// Show horizontally on the bottom of the viewer window
        /// </summary>
        Bottom
    }

    /// <summary>
    /// The ViewerMode when the window is being shown (i.e. not when it's hidden while idle)
    /// </summary>
    private ViewerMode _nominalViewerMode = ViewerMode.Normal;

    /// <summary>
    /// THe VisualTheme used to display prompted text
    /// </summary>
    private VisualTheme _visualTheme = VisualTheme.Light;

    /// <summary>
    /// Whether and where to show a progress bug on the prompt text viewer screen
    /// </summary>
    private ProgressBug _progressBug = ProgressBug.None;

    /// <summary>
    /// The brush to use for drawing text
    /// </summary>
    private SolidColorBrush _textForecolour = Brushes.Black;

    /// <summary>
    /// Sentinel used to indicate whether an action has occured that requires redrawing the prompted text
    /// </summary>
    private bool _redrawRequired = false;

    /// <summary>
    /// Sentinel used to override the default behaviour of canceling attempts to close the viewer window
    /// </summary>
    private bool _reallyCloseThisTime = false;

    /// <summary>
    /// The content being prompted as represented by the class which parses and interprets the content
    /// </summary>
    private readonly DvarContent _dvarContent = new();

    /// <summary>
    /// The content to be prompted
    /// </summary>
    internal DvarContent DvarContent { get { return _dvarContent; } }

    /// <summary>
    /// The index into the List<LineOfText> for the first line being drawn on the screen
    /// </summary>
    private int _firstLineOfTextIndex = 0;

    /// <summary>
    /// The number of pixels per density independent pixel for the screen on which this window is displayed.
    /// </summary>
    private readonly double _pixelsPerDip;

    /// <summary>
    /// Set the visual theme of the viewer window
    /// </summary>
    /// <param name="visualTheme">Enum specifying the supported visual theme to be applied</param>
    public void SetVisualTheme(VisualTheme visualTheme)
    {
        // Is the visual theme actually changing?
        VisualTheme oldTheme = _visualTheme;

        if (oldTheme != visualTheme)
        {
            _visualTheme = visualTheme;
            switch (visualTheme)
            {
                case VisualTheme.Light:
                    _textForecolour = Brushes.Black;
                    ContainerGrid.Background = Brushes.White;
                    ProgressBugRectangle.Fill = new SolidColorBrush(Color.FromArgb(0x20, 0x00, 0x00, 0x00));
                    this.Background = Brushes.White;
                    break;
                case VisualTheme.Dark:
                    _textForecolour = Brushes.White;
                    ContainerGrid.Background = Brushes.Black;
                    ProgressBugRectangle.Fill = new SolidColorBrush(Color.FromArgb(0x20, 0xff, 0xff, 0xff));
                    this.Background = Brushes.Black;
                    break;
                case VisualTheme.Matrix:
                    _textForecolour = new SolidColorBrush(Color.FromArgb(0xff, 0x00, 0xe6, 0x00));
                    ContainerGrid.Background = new SolidColorBrush(Color.FromArgb(0xff, 0x10, 0x1e, 0x10));
                    ProgressBugRectangle.Fill = new SolidColorBrush(Color.FromArgb(0x20, 0x00, 0xe6, 0x00));
                    this.Background = new SolidColorBrush(Color.FromArgb(0xff, 0x10, 0x1e, 0x10)); ;
                    break;
                default:
                    break;
            }
            InvalidateVisual();
        }
    }

    /// <summary>
    /// Set the visual presentation of the progress bug displayed (or not) in the viewer window
    /// </summary>
    /// <param name="progressBug">Enum indicating the desired location (if any) of the progress bug</param>
    public void SetProgressBug(ProgressBug progressBug)
    {
        _progressBug = progressBug;
        this.InvalidateVisual();
        if (progressBug == ProgressBug.None)
        {
            ProgressBugRectangle.Visibility = Visibility.Hidden;
        }
        else
        {
            ProgressBugRectangle.Visibility = Visibility.Visible;
        }
    }

    /// <summary>
    /// Loads content to be prompted from a file of a supported file type.
    /// </summary>
    /// <param name="filename">Full path and filespec of the file from which to extract content</param>
    public void LoadContentFromFile(string filename)
    {
        _dvarContent.LoadContentFromFile(filename);
        TriggerLayoutRecalculation();
        _firstLineOfTextIndex = 0;
        this.InvalidateVisual();
    }

    /// <summary>
    /// Recalculates the parsing of content fragments into lines based on current visual parameters
    /// </summary>
    private void TriggerLayoutRecalculation()
    {
        // Sanity check.
        int firstLineStartingFragmentIdx = 0;
        if (_dvarContent.LinesOfText.Count > 0)
        {
            // Before laying out the lines again. Bookmark the starting point so that it can be restored (roughly)
            firstLineStartingFragmentIdx = _dvarContent.LinesOfText[_firstLineOfTextIndex].Start;
        }

        _dvarContent.LayOutLines(
            new FontFamily(Typeface),
            FontSizeEms,
            _pixelsPerDip,
            AvailableWidth());

        if (_dvarContent.LinesOfText.Count > 0)
        {
            // Now find the (new) line which contains the bookmarked starting fragment index
            // and make that the new current scrolling location for painting prompt text
            int left = 0;
            int right = _dvarContent.LinesOfText.Count - 1;
            while (left <= right)
            {
                int mid = left + (right - left) / 2;
                // Is this the line? (bookmarked fragment is in the [mid] line)
                if ((_dvarContent.LinesOfText[mid].Start <= firstLineStartingFragmentIdx)
                    && ((_dvarContent.LinesOfText[mid].Start + _dvarContent.LinesOfText[mid].Count)
                       >= firstLineStartingFragmentIdx))
                {
                    _firstLineOfTextIndex = mid;
                    break;
                }
                else if (_dvarContent.LinesOfText[mid].Start > firstLineStartingFragmentIdx)
                {
                    // This line starts later than the current [mid] line (so shift to the right)
                    left = mid + 1;
                }
                else
                {
                    // The line must be before the current [mid] line (so shift to the left)
                    right = mid - 1;
                }
            }
        } 
    }

    /// <summary>
    /// The full path and filespec of the file from which the current content was extracted
    /// </summary>
    public string PrompterContentFilespec { get => _dvarContent.RawContentFilespec; }

    /// <summary>
    /// Change the window to either restored down/movable mode or fullscreen mode
    /// </summary>
    /// <param name="mode">The mode to assume</param>
    public void SetViewerMode(ViewerMode mode)
    {
        switch (mode)
        {
            case ViewerMode.Normal:
                this.ResizeMode = ResizeMode.CanResize;
                this.WindowStyle = WindowStyle.SingleBorderWindow;
                this.WindowState = WindowState.Normal;
                this.Topmost = false;

                _nominalViewerMode = mode;

                if (_redrawRequired)
                {
                    TriggerLayoutRecalculation();
                    this.InvalidateVisual();
                }

                break;
            case ViewerMode.Fullscreen:
                this.ResizeMode = ResizeMode.NoResize;
                this.WindowStyle = WindowStyle.None;
                this.WindowState = WindowState.Maximized;
                this.Topmost = true;

                if (_redrawRequired)
                {
                    TriggerLayoutRecalculation();
                    this.InvalidateVisual();
                }

                _nominalViewerMode = mode;

                break;
            case ViewerMode.Hidden:
                this.WindowState = WindowState.Minimized;
                this.ResizeMode = ResizeMode.NoResize;
                this.WindowStyle = WindowStyle.SingleBorderWindow;
                this.Topmost = false;

                // Don't set the nominal viewer mode.
                // Hidden is always a special/conditional case.
                break;
            default:
                break;
        }
    }

    /// <summary>
    /// Call to explicitly close the viewer window (manual attempts are rejected by default)
    /// </summary>
    public void ShutDownViewer()
    {
        _reallyCloseThisTime = true;
        this.Close();
    }

    private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        if (!_reallyCloseThisTime)
        {
            // Don't allow the window to close on its own.
            // The parent window (MainWindow) must close it.
            e.Cancel = true;
        }
    }

    /// <summary>
    /// The amount of usable space for drawing prompt text
    /// </summary>
    /// <returns>The available space in pixels</returns>
    private double AvailableWidth()
    {
        return this.Width
               - ContainerGrid.Margin.Left - ContainerGrid.Margin.Right
               - (2.0 * HorizontalMargin);
    }

    /// <summary>
    /// The number of lines of prompted text that fit within the available space.
    /// </summary>
    /// <returns>The number of lines that will fit completely (nearest integer value)</returns>
    private int LinesPerPage()
    {
        // Text starts appearing half a line down so allow a vertical margin of 
        // FontSizeEms above and below the usable space for painting.
        // Just account for the upper margin, however since the lower margin
        // is taken up by line spacing
        return (int)Math.Floor(
            (this.ActualHeight - ContainerGrid.Margin.Top - ContainerGrid.Margin.Bottom - FontSizeEms) 
            / LineHeight());
    }

    /// <summary>
    /// The height of a line of text taking into consideration the font size and line spacing
    /// </summary>
    /// <returns>The height of a line, including spacing, in pixels</returns>
    private double LineHeight()
    {
        return FontSizeEms * LineSpacing;
    }

    /// <summary>
    /// The size of the font for prompted text in Em units
    /// </summary>
    public double FontSizeEms
    {
        get => _fontSizeEms;
        set
        {
            _fontSizeEms = value;
            if (_redrawRequired)
            {
                TriggerLayoutRecalculation();
                this.InvalidateVisual();
            }
        }
    }

    /// <summary>
    /// The name of the _typeface to use for prompting text (must be a font family name that is installed on the computer)
    /// </summary>
    public string Typeface
    {
        get => _typeface;
        set
        {
            _typeface = value;
            if (_redrawRequired)
            {
                TriggerLayoutRecalculation();
                this.InvalidateVisual();
            }
        }
    }

    /// <summary>
    /// The space between lines as a factor of the FontSizeEms property (1.0 is single spacing)
    /// </summary>
    public double LineSpacing
    {
        get => lineSpacing;
        set
        { 
            lineSpacing = value;
            if (_redrawRequired)
            {
                TriggerLayoutRecalculation();
                this.InvalidateVisual();
            }
        }
    }

    /// <summary>
    /// The space to the left and right of the prompted text in device independent pixels
    /// </summary>
    public double HorizontalMargin
    {
        get => horizontalMargin;
        set
        { 
            horizontalMargin = value;
            if (_redrawRequired)
            {
                TriggerLayoutRecalculation();
                this.InvalidateVisual();
            }
        }
    }

    /// <summary>
    /// The buffer for the image of rendered prompt text
    /// </summary>
    private RenderTargetBitmap? _buffer;

    // Backing fields for properties that require actions on setting
    // -------------------------------------------------------------

    /// <summary>
    /// Backing field for the size of the font for prompted text in Em units
    /// </summary>
    private double _fontSizeEms = 32.0;

    /// <summary>
    /// Backing field for the name of the _typeface to use for prompting text (must be a font family name that is installed on the computer)
    /// </summary>
    private string _typeface = "Verdana";

    /// <summary>
    /// Backing field for the space to the left and right of the prompted text in device independent pixels
    /// </summary>
    private double horizontalMargin = 24.0;

    /// <summary>
    /// Backing field for the space between lines as a factor of the FontSizeEms property (1.0 is single spacing)
    /// </summary>
    private double lineSpacing = 2.0;

    /// <summary>
    /// The distance scrolled per unit time (negative values scroll text backwards)
    /// </summary>
    public double ScrollVelocity { get; set; } = 0.0;

    /// <summary>
    /// The visual context for the control where rendered prompt text is drawn
    /// </summary>
    private readonly DrawingVisual _drawingVisual = new();

    /// <summary>
    /// Timer used to control scrolling of prompted text
    /// </summary>
    private readonly DispatcherTimer _heartbeat;

    public ViewerWindow()
    {
        InitializeComponent();

        // Establish the screen properties...
        _pixelsPerDip = VisualTreeHelper.GetDpi(this).PixelsPerDip;

        // Set up the heartbeat time that watches for incoming HTTP requests
        _heartbeat = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(1)
        };
        _heartbeat.Tick += Heartbeat_Tick;
    }

    protected override void OnRender(DrawingContext drawingContext)
    {
        base.OnRender(drawingContext);

        _buffer = new RenderTargetBitmap((int)this.Prompter1Canvas.ActualWidth, (int)this.Prompter1Canvas.ActualHeight, 96, 96, PixelFormats.Pbgra32);
        Prompter1Image.Source = _buffer;
        DrawText();
    }

    /// <summary>
    /// Sets the timer interval (usually based on application configuration settings)
    /// </summary>
    /// <param name="milliseconds">The number of milliseconds between ticks</param>
    internal void SetTimerTickInterval(int milliseconds)
    {
        bool timerRunning = _heartbeat.IsEnabled;

        if (timerRunning)
        {
            _heartbeat.Stop();
        }
        _heartbeat.Interval = TimeSpan.FromMilliseconds(milliseconds);
        if (timerRunning)
        {
            _heartbeat.Start();
        }
    }

#if DEBUG
    private static void LogDrawTextAction(string logText, bool cleanFile = false)
    {
        string filespec = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "dvardeputydrawtext.log");
        if (cleanFile)
        {
            File.WriteAllText(filespec, logText + Environment.NewLine);
        }
        else
        {
            File.AppendAllText(filespec, logText + Environment.NewLine);
        }
    }
#endif

    /// <summary>
    /// Draw a page full of lines based on the current starting line index and the current lines per page
    /// </summary>
    private void DrawText()
    {
        // Sanity check
        if (_buffer is null) return;

        // Note the line height
        double lineHeight = LineHeight();
        
        // Set the starting point for drawing text
        // Start one line of text's height down 
        double verticalLineOffset = FontSizeEms;
        double horizontalFragmentOffset;
        double pxPerDip = VisualTreeHelper.GetDpi(this).PixelsPerDip;
        var fontFamily = new FontFamily(Typeface);

        // Get the starting and finishing lines but make sure we don't try to scroll out of bounds.
        int startingLineIndex = Math.Max(Math.Min(_firstLineOfTextIndex, _dvarContent.LinesOfText.Count - 1), 0);
        int endingLineIndex = Math.Min((LinesPerPage() + startingLineIndex - 1), (_dvarContent.LinesOfText.Count - 1));

#if DEBUG
        // Set the following to true for debugging.
        bool logDraw = false;
        if (logDraw)
        {
            LogDrawTextAction($"Start of DrawText: lines={startingLineIndex} to {endingLineIndex}", true);
            LogDrawTextAction($"AvailableWidth={AvailableWidth()}");
            LogDrawTextAction($"HorizontalMargin={HorizontalMargin}");
            LogDrawTextAction($"FontSizeEms={FontSizeEms}");
            LogDrawTextAction($"verticalLineOffset={verticalLineOffset}");
        }
#endif

        using (DrawingContext drawingContext = _drawingVisual.RenderOpen())
        {
            for (int lineIdx = startingLineIndex; lineIdx <= endingLineIndex; lineIdx++)
            {
                LineOfText lineOfText = _dvarContent.LinesOfText[lineIdx];
#if DEBUG
                if (logDraw)
                {
                    LogDrawTextAction("---Line----" + Environment.NewLine +
                        $"lineIdx={lineIdx} from frag:{lineOfText.Start} for {lineOfText.Count}. Margin={lineOfText.Margin} TabStop={lineOfText.TabStop} RTL={lineOfText.IsRightToLeft}");
                }
#endif
                // Is this LTR or RTL?
                bool isRTL = lineOfText.IsRightToLeft;
                if (isRTL)
                {
                    // Reset the horizontal starting point of the line.
                    // Available Width subtracts 2 x H margin, so add one back in.
                    horizontalFragmentOffset = AvailableWidth() + HorizontalMargin - lineOfText.Margin ;
                }
                else
                {
                    horizontalFragmentOffset = HorizontalMargin + lineOfText.Margin;
                }
#if DEBUG
                if (logDraw)
                {
                    LogDrawTextAction($"horizontalFragmentOffset={horizontalFragmentOffset}");
                }
#endif
                for (int fragIdx = 0; fragIdx < lineOfText.Count; fragIdx++)
                {
                    TextFragment fragment = _dvarContent.Fragments[(lineOfText.Start + fragIdx)];
                    string fragmentContent;
                    if (isRTL)
                    {
                        fragmentContent = (fragment.SuppressTrailingSpace ? "" : " ") + fragment.Content;
                    }
                    else
                    {
                        fragmentContent = fragment.Content + (fragment.SuppressTrailingSpace ? "" : " ");
                    }
#if DEBUG
                    if (logDraw)
                    {
                        string crlf = fragment.IsLineBreak ? " CRLF" : string.Empty;
                        string empty = fragment.IsEmpty ? " {}" : string.Empty;
                        string ishe = fragment.HasHebrew ? " HE" : string.Empty;
                        string fmt = (fragment.Bold ? "B" : string.Empty)
                                   + (fragment.Italic ? "I" : string.Empty)
                                   + (fragment.Underline ? "U" : string.Empty);
                        LogDrawTextAction("---" + Environment.NewLine
                            + $"fragIdx={fragIdx} fragContent=[{fragmentContent}] (len={fragment.Content.Length}) {fmt}{crlf}{empty}{ishe}");
                    }
#endif
                    // Create the formatted text for this fragment
                    FontStyle fontStyle = (fragment.Italic ? FontStyles.Italic : FontStyles.Normal);
                    FontWeight fontWeight = (fragment.Bold ? FontWeights.Bold : FontWeights.Normal);
                    var formattedText = new FormattedText(
                        fragmentContent,
                        CultureInfo.InvariantCulture,
                        FlowDirection.LeftToRight,
                        new Typeface(fontFamily, fontStyle, fontWeight, FontStretches.Normal),
                        FontSizeEms,
                        _textForecolour,
                        pxPerDip);

                    if (fragment.Underline)
                    {
                        TextDecorationCollection textDecorations = [];
                        textDecorations.Add(TextDecorations.Underline);
                        formattedText.SetTextDecorations(textDecorations);
                    }
                    // Calculate the correct fragment width based on being a regular
                    // fragment or a start of numbers/bulleted line...
                    double fragmentWidth = formattedText.WidthIncludingTrailingWhitespace;
#if DEBUG
                    if (logDraw)
                    {
                        LogDrawTextAction($"> Initial fragmentWidth={fragmentWidth}");
                    }
#endif
                    if (fragIdx == 0)
                    {
                        if (fragment.NumberingId > 0 && fragment.NumberingLevel >= 0)
                        {
                            fragmentWidth += lineOfText.TabStop;
#if DEBUG
                            if (logDraw)
                            {
                                LogDrawTextAction($"> Adjusted fragmentWidth={fragmentWidth} (for number/bullet)");
                            }
#endif
                        }
                    }    

                    // Offset for this fragment (RTL only)
                    if (isRTL)
                    {
                        horizontalFragmentOffset -= fragmentWidth;
#if DEBUG
                        if (logDraw)
                        {
                            LogDrawTextAction($"Adjusted horizontalFragmentOffset={horizontalFragmentOffset} for RTL fragment width");
                        }
#endif
                    }

                    // Draw the formatted text string to the DrawingContext of the control.
                    drawingContext.DrawText(formattedText, new Point(horizontalFragmentOffset, verticalLineOffset));
#if DEBUG
                    if (logDraw)
                    {
                        LogDrawTextAction($"Draw text at {horizontalFragmentOffset},{verticalLineOffset}");
                    }
#endif
                    // Offset for the next fragment (LTR only)
                    if (!isRTL)
                    {
                        horizontalFragmentOffset += fragmentWidth;
#if DEBUG
                        if (logDraw)
                        {
                            LogDrawTextAction($"Adjusted horizontalFragmentOffset={horizontalFragmentOffset} for LTR fragment width");
                        }
#endif
                    }
                }
                // Offset for the next line
                verticalLineOffset += lineHeight;
#if DEBUG
                if (logDraw)
                {
                    LogDrawTextAction($"verticalLineOffset={verticalLineOffset} (End of Line)");
                }
#endif
            }

            if (_progressBug != ProgressBug.None)
            {
                // Calculate the size and position of the progress bug
                double bugThickness = FontSizeEms * 0.66;
                switch (_progressBug)
                {
                    case ProgressBug.RightSide:
                        ProgressBugRectangle.Height = Math.Max(0, ((double)(endingLineIndex - startingLineIndex) / (double)_dvarContent.LinesOfText.Count) * Prompter1Image.ActualHeight);
                        ProgressBugRectangle.Width = bugThickness;
                        ProgressBugRectangle.SetValue(Canvas.LeftProperty, Math.Max(0, Prompter1Image.ActualWidth - bugThickness));
                        ProgressBugRectangle.SetValue(Canvas.TopProperty, Math.Max(0, ((double)startingLineIndex / (double)_dvarContent.LinesOfText.Count) * Prompter1Image.ActualHeight));
                        break;
                    case ProgressBug.Bottom:
                        ProgressBugRectangle.Height = bugThickness;
                        ProgressBugRectangle.Width = Math.Max(0, ((double)(endingLineIndex - startingLineIndex) / (double)_dvarContent.LinesOfText.Count) * Prompter1Image.ActualWidth);
                        ProgressBugRectangle.SetValue(Canvas.LeftProperty, Math.Max(0, ((double)startingLineIndex / (double)_dvarContent.LinesOfText.Count) * Prompter1Image.ActualWidth));
                        ProgressBugRectangle.SetValue(Canvas.TopProperty, Math.Max(0, Prompter1Image.ActualHeight - bugThickness));
                        break;
                    default:    // Including ProgressBug.None
                        // Won't happen.
                        break;
                }
            }
        }

        _buffer.Render(_drawingVisual);

        // Once the visual has been rendered the first time, any other invalidation requires redrawing.
        _redrawRequired = true;
    }

    private void Heartbeat_Tick(object? sender, EventArgs e)
    {
        Thickness imageMargin = this.Prompter1Image.Margin;

        if (ScrollVelocity >= 0)
        {
            imageMargin.Top -= ScrollVelocity;
            this.Prompter1Image.Margin = imageMargin;

            if (Math.Abs(imageMargin.Top) >= LineHeight())
            {
                // Scroll down one line
                _firstLineOfTextIndex++;

                imageMargin.Top = 0;
                this.Prompter1Image.Margin = imageMargin;

                this.InvalidateVisual();

                if (_firstLineOfTextIndex >= _dvarContent.LinesOfText.Count - LinesPerPage())
                {
                    StopScrolling();
                }
            }
        }
        else
        {
            imageMargin.Top += Math.Abs(ScrollVelocity);
            this.Prompter1Image.Margin = imageMargin;

            if (Math.Abs(imageMargin.Top) >= LineHeight())
            {
                // Scroll up one line
                _firstLineOfTextIndex--;

                imageMargin.Top = 0;
                this.Prompter1Image.Margin = imageMargin;

                this.InvalidateVisual();

                if (_firstLineOfTextIndex == 0)
                {
                    StopScrolling();
                }
            }
        }
    }

    /// <summary>
    /// Start the text moving according to the current ScrollVelocity
    /// </summary>
    public void StartScrolling()
    {
        _heartbeat.Start();
    }

    /// <summary>
    /// Stop the text from moving according to the current ScrollVelocity
    /// </summary>
    public void StopScrolling()
    { 
        _heartbeat.Stop(); 
    }

    /// <summary>
    /// Scroll down by one full page, stopping at the last page
    /// </summary>
    public void PageDown()
    {
        int linesPerPage = LinesPerPage();

        _firstLineOfTextIndex = 
            Math.Max(
                Math.Min(
                    _firstLineOfTextIndex + linesPerPage - 1,   // Don't go a full page, go 1 line less than a full page
                    _dvarContent.LinesOfText.Count - linesPerPage //+ 1
                    ),
                0
                );
        // Reset the scrolling position of the propmting image
        Thickness imageMargin = this.Prompter1Image.Margin;
        imageMargin.Top = 0;
        this.Prompter1Image.Margin = imageMargin;

        this.InvalidateVisual();
    }

    /// <summary>
    /// Scroll up by one full page, stopping at the first page
    /// </summary>
    public void PageUp()
    {
        int linesPerPage = LinesPerPage();

        _firstLineOfTextIndex =
        Math.Max(
                _firstLineOfTextIndex - linesPerPage + 1,   // Don't go a full page, go 1 line less than a full page
                0
                );
        // Reset the scrolling position of the propmting image
        Thickness imageMargin = this.Prompter1Image.Margin;
        imageMargin.Top = 0;
        this.Prompter1Image.Margin = imageMargin;

        this.InvalidateVisual();
    }

    /// <summary>
    /// Jump to the top of the content
    /// </summary>
    public void PageFirst()
    {
        _firstLineOfTextIndex = 0;
        // Reset the scrolling position of the propmting image
        Thickness imageMargin = this.Prompter1Image.Margin;
        imageMargin.Top = 0;
        this.Prompter1Image.Margin = imageMargin;

        this.InvalidateVisual();
    }

    /// <summary>
    /// Whether or not the text is scrolling at the current ScrollVelocity
    /// Note that if the velocity is 0.0, the text may still be considered to be scrolling.
    /// </summary>
    /// <returns>True if the scrolling timer is active, false otherwise</returns>
    public bool IsScrolling()
    {
        return _heartbeat.IsEnabled;
    }

    private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        if (_redrawRequired)
        {
            TriggerLayoutRecalculation();
            this.InvalidateVisual();
        }
    }
}
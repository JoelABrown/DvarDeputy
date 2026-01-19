using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Mooseware.Tachnit.AtemApi;
using Mooseware.DvarDeputy.Controls;
using System.Collections.Concurrent;
using System.Drawing.Text;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Animation;
using System.Windows.Threading;

namespace Mooseware.DvarDeputy;

// TODO: NEXT: Figure out if there's a way to track the starting fragment index so that a rapid invocation of font or spacing or
//             margin changes (or font type?) don't miss the point of freezing the fragment idx 
// TODO: When numbering schemes have a starting number take that into account (maybe a version X feature?)
// TODO: Determine if fancy background painting is required of if just painting in chunks in the main UI thread is good enough
// TODO: Follow up in these URLs if fancier background painting is required:
//       - https://learn.microsoft.com/en-us/dotnet/desktop/wpf/graphics-multimedia/using-drawingvisual-objects
//       - https://learn.microsoft.com/en-us/dotnet/desktop/wpf/advanced/how-to-draw-text-to-a-visual
//       - https://learn.microsoft.com/en-us/dotnet/desktop/wpf/advanced/advanced-text-formatting
//       - https://learn.microsoft.com/en-us/dotnet/desktop/wpf/advanced/drawing-formatted-text
//       - https://github.com/Microsoft/WPF-Samples/tree/main/Visual%20Layer/DrawingVisual
//       - https://swharden.com/csdv/system.drawing/quickstart-wpf/
//       - https://www.gamedev.net/forums/topic/670737-how-can-i-draw-rectangles-on-a-wpf-canvas/
//       - https://stackoverflow.com/questions/452726/scrolling-credits-screen-in-wpf-ideas
//       - https://www.google.com/search?q=C%23+DrawingContext+render+in+background+thread&sca_esv=f581f0a2a024c0bb&sxsrf=AE3TifM_njwGhSYvquym49YInr7k0UXs8A%3A1766755316802&ei=9ItOab7dMIfx0PEPp4uduQQ&ved=0ahUKEwi-p8C4rNuRAxWHODQIHadFJ0cQ4dUDCBE&uact=5&oq=C%23+DrawingContext+render+in+background+thread&gs_lp=Egxnd3Mtd2l6LXNlcnAiLUMjIERyYXdpbmdDb250ZXh0IHJlbmRlciBpbiBiYWNrZ3JvdW5kIHRocmVhZDIFECEYoAFI-4sBUJkGWOaJAXAGeACQAQCYAWigAdAgqgEENDkuMrgBA8gBAPgBAZgCOaAC2SGoAhDCAgsQABiJBRiiBBiwA8ICBBAjGCfCAgoQIxiABBiKBRgnwgILEAAYgAQYigUYkQLCAgoQABiABBiKBRhDwgINEAAYgAQYigUYQxixA8ICCxAAGIAEGLEDGIMBwgIREC4YgAQYsQMYgwEYxwEY0QPCAgsQLhiABBixAxiDAcICERAAGIAEGIoFGI0GGLEDGIMBwgIHECMY6gIYJ8ICDRAjGIAEGIoFGOoCGCfCAhcQABiABBiKBRiRAhjnBhjqAhi0AtgBAcICEBAuGIAEGIoFGEMYxwEY0QPCAggQABiABBixA8ICBRAAGIAEwgIKEAAYgAQYFBiHAsICBhAAGBYYHsICBhAAGB4YDcICBRAAGO8FwgIIEAAYgAQYogTCAggQABiJBRiiBMICCBAAGAgYHhgNwgIEECEYFcICBxAhGAoYoAHCAgUQIRifBZgDA4gGAZAGAboGBggBEAEYAZIHBDUzLjSgB7u6AbIHBDQ3LjS4B84hwgcHMjMuMzEuM8gHU4AIAQ&sclient=gws-wiz-serp
//       - https://learn.microsoft.com/en-us/dotnet/api/system.windows.media.imaging.writeablebitmap?view=windowsdesktop-10.0

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{

    /// <summary>
    /// Host for listening to HTTP API calls (for example, from a Companion-powered Stream Deck)
    /// </summary>
    private readonly IHost _host;

    /// <summary>
    /// Timer for ticking through a loop to look for queued messages
    /// </summary>
    private readonly DispatcherTimer _heartbeat;

    /// <summary>
    /// The queue of messages being received via the HTTP API
    /// </summary>
    private readonly ConcurrentQueue<ApiMessage> _messageQueue;

    /// <summary>
    /// Reference to the prompt text viewing window
    /// </summary>
    private readonly ViewerWindow _viewerWindow = new();

    /// <summary>
    /// The ATEM Switcher controller which wraps the necessary parts of the ATEM API
    /// </summary>
    private AtemSwitcher? _atemSwitcher;

    /// <summary>
    /// The short name of the video input to restore in the ATEM switcher when the viewer window is hidden.
    /// </summary>
    private string _switcherInputToRestore = string.Empty;

    /// <summary>
    /// Sentinel to track when the window is loading so that initial settings of controls don't cause recursion
    /// </summary>
    private bool _loading = false;

    /// <summary>
    /// Base constructor for the MainWindow
    /// </summary>
    /// <param name="msgQueue">ConcurrentQueue<ApiMessage> for the HTTP API injected via DI</param>
    public MainWindow(ConcurrentQueue<ApiMessage> msgQueue)
    {
        InitializeComponent();
        _loading = true;
        _viewerWindow.Hide();

        // Set the local reference to the (singleton) ConcurrentQueue for the UI thread.
        _messageQueue = msgQueue;

        // Set up the heartbeat time that watches for incoming HTTP requests
        _heartbeat = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(100)
        };
        _heartbeat.Tick += Heartbeat_Tick;

        // Restore any settings from the previous session and load application settings...
        if (Properties.Settings.Default.UpgradeRequired)
        {
            Properties.Settings.Default.Upgrade();
            Properties.Settings.Default.UpgradeRequired = false;
            Properties.Settings.Default.Save();
        }
        Properties.Settings.Default.Reload();

        // Get a connection to the ATEM switcher
        string atemIpAddress = Properties.Settings.Default.AtemIpAddress;
        if (atemIpAddress is null || atemIpAddress.Length == 0)
        {
            atemIpAddress = "192.168.1.240";    // Apply a default
        }
        ConnectToAtem(atemIpAddress);


        //Get the base URL for .UseUrls() from the App.config settings file
        string webApiRootUrl = Properties.Settings.Default.WebApiUrlRoot;
        string webApiUrlPort = Properties.Settings.Default.WebApiUrlPort;

        // Create the background Web API server running within this WPF app...
        var builder = Host.CreateDefaultBuilder()
            .ConfigureWebHostDefaults(webBuilder =>
            {
                webBuilder.UseStartup<Startup>();
                webBuilder.UseUrls(webApiRootUrl + ":" + webApiUrlPort);
            });

        // Pass a reference to the ConcurrentQueue for the web server thread so that web APIs can add to the queue
        builder.ConfigureServices((hostContext, services) =>
        {
            services.AddSingleton<ConcurrentQueue<ApiMessage>>(_messageQueue);
        });
        _host = builder.Build();
        _host.Start();
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {

        FillFontsComboBox();

        // Apply application control settings (and their related user control settings)
        _viewerWindow.SetTimerTickInterval(Properties.Settings.Default.TimerFrequency);
        ScrollSpeedSlider.Minimum = Properties.Settings.Default.LowerScrollLimit;
        ScrollSpeedSlider.Maximum = Properties.Settings.Default.UpperScrollLimit;
        ScrollSpeedSlider.Value = Properties.Settings.Default.LastScrollVelocity;
        FontSizeSlider.Minimum = Properties.Settings.Default.LowerFontSizeLimit;
        FontSizeSlider.Maximum = Properties.Settings.Default.UpperFontSizeLimit;
        FontSizeSlider.Value = Properties.Settings.Default.LastFontSize;
        SideMarginSlider.Maximum = Properties.Settings.Default.UpperHorizontalMarginLimit;
        SideMarginSlider.Value = Properties.Settings.Default.LastHorizontalMargin;
        LineSpacingSlider.Maximum = Properties.Settings.Default.UpperLineSpacingLimit;
        LineSpacingSlider.Value = Properties.Settings.Default.LastLineSpacing;
        if (Enum.TryParse<ViewerWindow.VisualTheme>(Properties.Settings.Default.VisualTheme, out var visualTheme))
        {
            switch (visualTheme)
            {
                case ViewerWindow.VisualTheme.Light:
                    ThemeLightRadioButton.IsChecked = true;
                    break;
                case ViewerWindow.VisualTheme.Dark:
                    ThemeDarkRadioButton.IsChecked = true;
                    break;
                case ViewerWindow.VisualTheme.Matrix:
                    ThemeMatrixRadioButton.IsChecked = true;
                    break;
                default:
                    break;
            }
            _viewerWindow.SetVisualTheme(visualTheme);
        }
        else
        {
            ThemeLightRadioButton.IsChecked = true;
        }
        if (Enum.TryParse<ViewerWindow.ProgressBug>(Properties.Settings.Default.ProgressBug, out var progressBug))
        {
            switch (progressBug)
            {
                case ViewerWindow.ProgressBug.None:
                    ProgressBugNoneRadioButton.IsChecked = true;
                    break;
                case ViewerWindow.ProgressBug.RightSide:
                    ProgressBugSideRadioButton.IsChecked = true;
                    break;
                case ViewerWindow.ProgressBug.Bottom:
                    ProgressBugBottomRadioButton.IsChecked = true;
                    break;
                default:
                    break;
            }
            _viewerWindow.SetProgressBug(progressBug);
        }
        else
        {
            ProgressBugNoneRadioButton.IsChecked = true;
        }

        // Apply content and view user settings.
        RestoreWindowLocationsCheckbox.IsChecked = Properties.Settings.Default.RestoreWindowLocations;
        if (Properties.Settings.Default.RestoreWindowLocations)
        {
            this.Width = Properties.Settings.Default.ControlWindowSize.Width;
            this.Height = Properties.Settings.Default.ControlWindowSize.Height;
            this.Top = Properties.Settings.Default.ControlWindowLocation.Y;
            this.Left = Properties.Settings.Default.ControlWindowLocation.X;
            _viewerWindow.Width = Properties.Settings.Default.ViewerWindowSize.Width;
            _viewerWindow.Height = Properties.Settings.Default.ViewerWindowSize.Height;
            _viewerWindow.Top = Properties.Settings.Default.ViewerWindowLocation.Y;
            _viewerWindow.Left = Properties.Settings.Default.ViewerWindowLocation.X;
        }
        _viewerWindow.Show();
        _viewerWindow.SetViewerMode(ViewerWindow.ViewerMode.Normal);
        RestoreViewerRadioButton.IsChecked = true;

        ReopenLastContentCheckbox.IsChecked = Properties.Settings.Default.ReopenLastContent;
        if (Properties.Settings.Default.ReopenLastContent)
        {
            ContentFilespecTextBox.Text = Properties.Settings.Default.LastSourceFile;
            _viewerWindow.LoadContentFromFile(ContentFilespecTextBox.Text);
        }

        // Check the settings for the ATEM switcher connection
        AtemApiStatusLed.IsOn = true;
        if (_atemSwitcher is not null && _atemSwitcher.IsReady)
        {
            AtemIpAddressTextBox.Text = _atemSwitcher.IpAddress;
            AtemApiStatusLed.SelectedColour = LightEmittingDiode.ColourOptions.Green;
            AtemInputsComboBox.IsEnabled = true;
            // Fill the inputs combo box
            AtemInputsComboBox.ItemsSource = _atemSwitcher.GetInputShortNames(AtemSwitcherPortType.External);
            int selectedIdx = AtemInputsComboBox.Items.IndexOf(Properties.Settings.Default.ViewerAtemInput);
            if (selectedIdx >= 0)
            {
                AtemInputsComboBox.SelectedIndex = selectedIdx;
            }
            else
            {
                AtemInputsComboBox.SelectedIndex = 0;
            }
        }
        else
        {
            AtemApiStatusLed.SelectedColour = LightEmittingDiode.ColourOptions.Red;
            AtemInputsComboBox.IsEnabled = false;
            AtemIpAddressTextBox.Text = Properties.Settings.Default.AtemIpAddress ?? string.Empty;
        }

        // Show the current application version number and copyright notice.
        string copyrightNotice = string.Empty;
        var attribs = Assembly.GetEntryAssembly()!.GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
        if (attribs.Length > 0)
        {
            copyrightNotice = ((AssemblyCopyrightAttribute)attribs[0]).Copyright;
        }
        AppVersionInfoLabel.Content = "Version: " + Assembly.GetExecutingAssembly()!.GetName()!.Version!.ToString();
        CopyrightInfoLabel.Content = copyrightNotice;

        // Clear the development time content of the status message textblock
        FlashPlaybackMessage(string.Empty);

        _loading = false;

        // Start looking for API messages
        _heartbeat.Start();
    }

    private async void Window_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        _heartbeat.Stop();

        // See https://stackoverflow.com/questions/67674514/hosting-an-asp-web-api-inside-a-wpf-app-wont-stop-gracefully
        // for why this particular sequence of events are required to allow the _host to be shut down gracefully when
        // the WPF application is shut down.
        e.Cancel = true;
        await _host.StopAsync();
        _host.Dispose();

        // Save settings as appropriate...
        if (Properties.Settings.Default.RestoreWindowLocations)
        {
            Properties.Settings.Default.ControlWindowLocation
                = new System.Drawing.Point(Convert.ToInt32(this.Left), Convert.ToInt32(this.Top));
            Properties.Settings.Default.ControlWindowSize
                = new System.Drawing.Size(Convert.ToInt32(this.Width), Convert.ToInt32(this.Height));
            if (_viewerWindow.WindowState != WindowState.Minimized)
            {
                Properties.Settings.Default.ViewerWindowLocation
                    = new System.Drawing.Point(Convert.ToInt32(_viewerWindow.Left), Convert.ToInt32(_viewerWindow.Top));
                Properties.Settings.Default.ViewerWindowSize
                    = new System.Drawing.Size(Convert.ToInt32(_viewerWindow.Width), Convert.ToInt32(_viewerWindow.Height));
            }
        }

        Properties.Settings.Default.LastFontSize = _viewerWindow.FontSizeEms;
        Properties.Settings.Default.LastHorizontalMargin = _viewerWindow.HorizontalMargin;
        Properties.Settings.Default.LastLineSpacing = _viewerWindow.LineSpacing;
        Properties.Settings.Default.LastScrollVelocity = _viewerWindow.ScrollVelocity;
        Properties.Settings.Default.LastTypeface = _viewerWindow.Typeface;

        if (Properties.Settings.Default.ReopenLastContent)
        {
            Properties.Settings.Default.LastSourceFile = _viewerWindow.PrompterContentFilespec;
        }
        else
        {
            Properties.Settings.Default.LastSourceFile = string.Empty;
        }

        ViewerWindow.VisualTheme visualTheme = ViewerWindow.VisualTheme.Light;
        if (ThemeLightRadioButton.IsChecked == true) visualTheme = ViewerWindow.VisualTheme.Light;
        if (ThemeDarkRadioButton.IsChecked == true) visualTheme = ViewerWindow.VisualTheme.Dark;
        if (ThemeMatrixRadioButton.IsChecked == true) visualTheme = ViewerWindow.VisualTheme.Matrix;
        Properties.Settings.Default.VisualTheme = visualTheme.ToString();

        ViewerWindow.ProgressBug progressBug = ViewerWindow.ProgressBug.None;
        if (ProgressBugNoneRadioButton.IsChecked == true) progressBug = ViewerWindow.ProgressBug.None;
        if (ProgressBugSideRadioButton.IsChecked == true) progressBug = ViewerWindow.ProgressBug.RightSide;
        if (ProgressBugBottomRadioButton.IsChecked == true) progressBug = ViewerWindow.ProgressBug.Bottom;
        Properties.Settings.Default.ProgressBug = progressBug.ToString();

        Properties.Settings.Default.Save();

        _viewerWindow.ShutDownViewer();

        try
        {
            this.Closing -= Window_Closing;
            Close();
        }
        catch (Exception)
        {   // Sometimes timing gets out of whack.
            Thread.Sleep(200);
        }
    }

    private void StartScrollingButton_Click(object sender, RoutedEventArgs e)
    {
        StartScrollingForwards();
    }

    /// <summary>
    /// Starts the process of scrolling the prompt text in the forward direction
    /// </summary>
    private void StartScrollingForwards()
    {
        if (_viewerWindow.ScrollVelocity < 0)
        {
            // Reverse direction
            _viewerWindow.ScrollVelocity *= -1;
        }
        _viewerWindow.StartScrolling();
    }

    private void ScrollBackButton_Click(object sender, RoutedEventArgs e)
    {
        if (_viewerWindow.ScrollVelocity > 0)
        {
            // Reverse direction
            _viewerWindow.ScrollVelocity *= -1;
        }
        _viewerWindow.StartScrolling();
    }

    /// <summary>
    /// Starts the process of scrolling the prompt text in the reverse direction
    /// </summary>
    private void StartScrollingBackwards()
    {
        if (_viewerWindow.ScrollVelocity > 0)
        {
            // Reverse direction
            _viewerWindow.ScrollVelocity *= -1;
        }
        _viewerWindow.StartScrolling();
    }

    private void StopScrollingButton_Click(object sender, RoutedEventArgs e)
    {
        _viewerWindow.StopScrolling();
#if DEBUG
        DebugDumpFragmentsAndLines();
#endif
    }

#if DEBUG
    /// <summary>
    /// Debug code which dumps out the list of parsed fragments and the most recently parsed lines of text
    /// </summary>
    private void DebugDumpFragmentsAndLines()
    {
        StringBuilder sb = new();

        int i = 0;
        foreach (var item in _viewerWindow.DvarContent.Fragments)
        {
            string crlf = item.IsLineBreak ? " CRLF" : string.Empty;
            string empty = item.IsEmpty ? " {}" : string.Empty;
            string ishe = item.HasHebrew ? " HE" : string.Empty;
            string fmt = (item.Bold ? "B" : string.Empty)
                       + (item.Italic ? "I" : string.Empty)
                       + (item.Underline ? "U" : string.Empty);

            sb.AppendLine($"Frag {i}:[{item.Content}] (len={item.Content.Length}) {fmt}{crlf}{empty}{ishe}");
            i++;
        }
        i = 0;
        foreach (var item in _viewerWindow.DvarContent.LinesOfText)
        {
            string text = string.Empty;
            for (int j = item.Start; j < item.Start + item.Count; j++)
            {
                var frag = _viewerWindow.DvarContent.Fragments[j];
                if (frag.SuppressTrailingSpace)
                {
                    text += frag.Content;
                }
                else
                {
                    text += frag.Content + " ";
                }
                if (frag.IsLineBreak)
                {
                    text += "[CRLF]";
                }
            }
            string rtl = string.Empty;
            double margin = item.Margin;
            double tabstop = item.TabStop;

            rtl = item.IsRightToLeft ? "RTL " : string.Empty;
            sb.AppendLine($"Line {i}: {item.Start} for {item.Count} = [{text}] {rtl} m={margin} t={tabstop}");
            i++;
        }

        File.WriteAllText(
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "dvardeputyfragline.log"), sb.ToString());
    }
#endif

    private void OpenWordDocumentButton_Click(object sender, RoutedEventArgs e)
    {
        Microsoft.Win32.OpenFileDialog dlg = new()
        {
            AddExtension = true,
            CheckFileExists = true,
            CheckPathExists = true,
            DefaultExt = ".docx",
            Filter = "Word Documents|*.docx|Text Files|*.txt",
            Multiselect = false,
            Title = "D'var Deputy - Open Content Document"
        };
        if (dlg.ShowDialog() ?? false)
        {
            if (dlg.FileName is not null && File.Exists(dlg.FileName))
            {
                ContentFilespecTextBox.Text = dlg.FileName;
                _viewerWindow.LoadContentFromFile(dlg.FileName);
            }
        }
    }

    /// <summary>
    /// Populates the FontFamilyComboBox with installed fonts
    /// </summary>
    private void FillFontsComboBox()
    {
        // Clear the combo box of any existing data
        FontFamilyComboBox.Items.Clear();
        // Add a prompting entry
        FontFamilyComboBox.Items.Add("(choose)");
        // Read the list of installed fonts and add them to the combobox
        using InstalledFontCollection fontsCollection = new();
        System.Drawing.FontFamily[] fontFamilies = fontsCollection.Families;
        foreach (System.Drawing.FontFamily font in fontFamilies)
        {
            FontFamilyComboBox.Items.Add(font.Name);
        }
        // Make an initial selection
        string lastUsedTypeface = Properties.Settings.Default.LastTypeface;
        if (FontFamilyComboBox.Items.Contains(lastUsedTypeface))
        {
            FontFamilyComboBox.SelectedItem = lastUsedTypeface;
        }
        else
        {
            FontFamilyComboBox.SelectedIndex = 0;
        }
    }

    private void RestoreWindowLocationsCheckbox_Checked(object sender, RoutedEventArgs e)
    {
        if (!_loading)
        {
            Properties.Settings.Default.RestoreWindowLocations = (bool)(RestoreWindowLocationsCheckbox.IsChecked ?? false);
        }
    }

    private void ReopenLastContentCheckbox_Checked(object sender, RoutedEventArgs e)
    {
        if (!_loading)
        {
            Properties.Settings.Default.ReopenLastContent = (bool)(ReopenLastContentCheckbox.IsChecked ?? false);
        }
    }

    private void FontSizeSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        int normalizedValue = (int)e.NewValue;
        FontSizeTextBlock.Text = normalizedValue.ToString() + " em";
        // Apply this value to the viewer window
        _viewerWindow.FontSizeEms = normalizedValue;
    }

    private void SideMarginSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        int normalizedValue = (int)e.NewValue;
        SideMarginTextBlock.Text = normalizedValue.ToString() + " px";
        // Apply this value to the viewer window
        _viewerWindow.HorizontalMargin = (double)normalizedValue;
    }

    private void LineSpacingSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        int normalizedMantissa = (int)(e.NewValue * 100);
        double normalizedValue = (double)normalizedMantissa / 100.0;
        LineSpacingTextBlock.Text = normalizedValue.ToString("0.00");
        // Apply this value to the viewer window
        _viewerWindow.LineSpacing = normalizedValue;
    }

    private void ScrollSpeedSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        int normalizedMantissa = (int)(e.NewValue * 100);
        double normalizedValue = (double)normalizedMantissa / 100.0;
        ScrollSpeedTextBlock.Text = normalizedValue.ToString("0.00");
        // Apply this value to the viewer window
        _viewerWindow.ScrollVelocity = normalizedValue;

    }

    private void FontFamilyComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        // Apply this value to the viewer window
        if (FontFamilyComboBox.SelectedIndex > 0)
        {
            System.Windows.Media.FontFamily newFont = new(FontFamilyComboBox.SelectedItem.ToString());
            _viewerWindow.Typeface = newFont.Source;
        }
    }

    private void PageForwardButton_Click(object sender, RoutedEventArgs e)
    {
        _viewerWindow.PageDown();
    }

    private void PageBackButton_Click(object sender, RoutedEventArgs e)
    {
        _viewerWindow.PageUp();
    }

    private void FontSizeResetButton_Click(object sender, RoutedEventArgs e)
    {
        FontSizeSlider.Value = Properties.Settings.Default.DefaultFontSize;
    }

    private void SideMarginResetButton_Click(object sender, RoutedEventArgs e)
    {
        SideMarginSlider.Value = Properties.Settings.Default.DefaultHorizontalMargin;
    }

    private void LineSpacingResetButton_Click(object sender, RoutedEventArgs e)
    {
        LineSpacingSlider.Value = Properties.Settings.Default.DefaultLineSpacing;
    }

    private void ScrollSpeedResetButton_Click(object sender, RoutedEventArgs e)
    {
        ScrollSpeedSlider.Value = Properties.Settings.Default.DefaultScrollVelocity;
    }

    private void ViewerViewModeOption_Click(object sender, RoutedEventArgs e)
    {
        // Which option was selected?
        RadioButton? radioButton = sender as RadioButton;
        if (radioButton is not null && radioButton.IsChecked is not null && (bool)radioButton.IsChecked)
        {
            switch (radioButton!.Name)
            {
                case "RestoreViewerRadioButton":
                    _viewerWindow.SetViewerMode(ViewerWindow.ViewerMode.Normal);
                    if (ViewerShownCheckBox is not null)
                    {
                        ViewerShownCheckBox.IsChecked = null;
                        ViewerShownCheckBox.IsEnabled = false;
                        ViewerShownCheckBoxStateChanged();
                    }
                    break;
                case "MaximizeViewerRadioButton":
                    _viewerWindow.SetViewerMode(ViewerWindow.ViewerMode.Fullscreen);
                    if (ViewerShownCheckBox is not null)
                    {
                        ViewerShownCheckBox.IsChecked = Properties.Settings.Default.ShowViewerWhenMaximized;
                        ViewerShownCheckBox.IsEnabled = true;
                    }
                    break;
                case "MinimizeViewerRadioButton":
                    _viewerWindow.SetViewerMode(ViewerWindow.ViewerMode.Hidden);
                    if (ViewerShownCheckBox is not null)
                    {
                        ViewerShownCheckBox.IsChecked = null;
                        ViewerShownCheckBox.IsEnabled = false;
                        ViewerShownCheckBoxStateChanged();
                    }
                    break;
                default:
                    break;
            }
        }
    }

    private void ViewerShownCheckBox_Click(object sender, RoutedEventArgs e)
    {
        ViewerShownCheckBoxStateChanged();
    }

    private void ViewerShownCheckBox_Checked(object sender, RoutedEventArgs e)
    {
        ViewerShownCheckBoxStateChanged();
    }

    /// <summary>
    /// Handles changes (however they may have been triggered) to the ViewerShownCheckBox
    /// </summary>
    private void ViewerShownCheckBoxStateChanged()
    { 
        if (ViewerShownCheckBox.IsChecked is null)
        {
            // Don't show the viewer in the ATEM Aux Out
            SetViewerRouting(false);
        }
        else
        {
            // Make a note for next time.
            Properties.Settings.Default.ShowViewerWhenMaximized = (bool)ViewerShownCheckBox.IsChecked;

            // Show or hide the ViewerWindow in the ATEM depending on the current selection
            if (ViewerShownCheckBox.IsChecked == true)
            {
                // Show it
                SetViewerRouting(true);
            }
            else
            {
                // Hide it
                SetViewerRouting(false);
            }
        }
    }

    /// <summary>
    /// Sets the routing of the ViewerWindow display in the ATEM switcher according to settings
    /// </summary>
    /// <param name="show">Use True to show the Viewer via the ATEM or False to hide it and restore the original video source</param>
    private void SetViewerRouting(bool show)
    {
        if (_atemSwitcher is not null)
        {
            if (show)
            {
                // Note what input is being routed so that it can be restored when the viewer is hidden.
                switch (Properties.Settings.Default.ViewerAtemOutput.ToUpper().Trim())
                {
                    case "AUX":
                        _switcherInputToRestore = _atemSwitcher.GetAuxInput().ShortName;
                        _atemSwitcher.SetAuxInput(Properties.Settings.Default.ViewerAtemInput);
                        break;
                    case "PGM":
                        _switcherInputToRestore = _atemSwitcher.GetProgramInput().ShortName;
                        _atemSwitcher.SetProgramInput(Properties.Settings.Default.ViewerAtemInput);
                        break;
                    case "PVW":
                        _switcherInputToRestore = _atemSwitcher.GetPreviewInput().ShortName;
                        _atemSwitcher.SetPreviewInput(Properties.Settings.Default.ViewerAtemInput);
                        break;
                    default:
                        // Nothing to do. This is an unknown output
                        break;
                }
            }
            else
            {
                switch (Properties.Settings.Default.ViewerAtemOutput.ToUpper().Trim())
                {
                    case "AUX":
                        _atemSwitcher.SetAuxInput(_switcherInputToRestore);
                        break;
                    case "PGM":
                        _atemSwitcher.SetProgramInput(_switcherInputToRestore);
                        break;
                    case "PVW":
                        _atemSwitcher.SetPreviewInput(_switcherInputToRestore);
                        break;
                    default:
                        // Nothing to do. This is an unknown output
                        break;
                }
            }
        }
    }

    private void ContentFileGrid_DragOver(object sender, DragEventArgs e)
    {
        if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            e.Effects = DragDropEffects.Copy;
        }
        else
        {
            e.Effects = DragDropEffects.None;
        }
    }

    private void ContentFileGrid_Drop(object sender, DragEventArgs e)
    {
        if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            // Get the list of file(s) being dropped...
            string[] droppedFiles = (string[])e.Data.GetData(DataFormats.FileDrop);

            // We can only load one, so pick the first one.
            if (droppedFiles.Length > 0)
            {
                if (droppedFiles[0] is not null && File.Exists(droppedFiles[0]))
                {
                    ContentFilespecTextBox.Text = droppedFiles[0];
                    _viewerWindow.LoadContentFromFile(droppedFiles[0]);
                }
            }
        }
    }

    private void Heartbeat_Tick(object? sender, EventArgs e)
    {
        // Look for API messages arrived since the last cycle...
        while (!_messageQueue.IsEmpty)
        {
            if (_messageQueue.TryDequeue(out ApiMessage? queueItem))
            {
                if (queueItem is not null)
                {
                    HandleApiCommand(queueItem);
                }
            }
        }
    }

    /// <summary>
    /// Process the commands received through the REST API 
    /// </summary>
    /// <param name="message">The API message details which specify what command to execute</param>
    private void HandleApiCommand(ApiMessage message)
    {
        string? flashMessage;
        double newScalarValue;

        // Make a message to flash to the UI so the operator knows that things are happening in the background
        flashMessage = message.Verb.ToString() + ": (" + message.Parameters + ")";

        // Do the actions requested by the API...
        switch (message.Verb)
        {
            case ApiMessageVerb.None:
                break;
            case ApiMessageVerb.Viewer:
                switch (message.Parameters)
                {
                    case ApiMessage.ViewerShow:
                        flashMessage = "Show viewer window";
                        MaximizeViewerRadioButton.IsChecked = true;
                        ViewerShownCheckBox.IsChecked = true;
                        break;
                    case ApiMessage.ViewerHide:
                        flashMessage = "Hide viewer window";
                        MinimizeViewerRadioButton.IsChecked = true;
                        break;
                    default:
                        break;
                }
                break;
            case ApiMessageVerb.Page:
                switch (message.Parameters)
                {
                    case ApiMessage.PageNext:
                        flashMessage = "Page Forward";
                        _viewerWindow.PageDown();
                        break;
                    case ApiMessage.PagePrevious:
                        flashMessage = "Page Back";
                        _viewerWindow.PageUp();
                        break;
                    case ApiMessage.PageFirst:
                        flashMessage = "Go to First Page";
                        _viewerWindow.PageFirst();
                        break;
                    default:
                        break;
                }
                break;
            case ApiMessageVerb.Scroll:
                switch (message.Parameters)
                {
                    case ApiMessage.ScrollBackward:
                        flashMessage = "Scroll Backwards";
                        StartScrollingBackwards();
                        break;
                    case ApiMessage.ScrollForward:
                        flashMessage = "Scroll Forwards";
                        StartScrollingForwards();
                        break;
                    case ApiMessage.ScrollStop:
                        flashMessage = "Stop Scrolling";
                        _viewerWindow.StopScrolling();
                        break;
                    default:
                        break;
                }
                break;
            case ApiMessageVerb.Font:
                switch (message.Parameters)
                {
                    case ApiMessage.FontIncrease:
                        // Can we increase the font by the requested amount?
                        // If so do it. Otherwise go as high as settings will allow.
                        newScalarValue = Math.Round(Math.Min(FontSizeSlider.Value + message.Scalar,
                            Properties.Settings.Default.UpperFontSizeLimit), 2);
                        FontSizeSlider.Value = newScalarValue;
                        flashMessage = "Increase font size to " + newScalarValue.ToString();
                        break;
                    case ApiMessage.FontDecrease:
                        // Can we decrease the font by the requested amount?
                        // If so do it. Otherwise go as low as settings will allow.
                        newScalarValue = Math.Round(Math.Max(FontSizeSlider.Value - message.Scalar,
                            Properties.Settings.Default.LowerFontSizeLimit), 2);
                        FontSizeSlider.Value = newScalarValue;
                        flashMessage = "Decrease font size to " + newScalarValue.ToString();
                        break;
                    case ApiMessage.FontReset:
                        FontSizeSlider.Value = Properties.Settings.Default.DefaultFontSize;
                        flashMessage = "Reset font size to default";
                        break;
                    default:
                        break;
                }
                break;
            case ApiMessageVerb.ScrollSpeed:
                switch (message.Parameters)
                {
                    case ApiMessage.ScrollSpeedIncrease:
                        // Can we increase the ScrollSpeed by the requested amount?
                        // If so do it. Otherwise go as high as settings will allow.
                        newScalarValue = Math.Round(Math.Min(ScrollSpeedSlider.Value + message.Scalar,
                            Properties.Settings.Default.UpperScrollLimit), 2);
                        ScrollSpeedSlider.Value = newScalarValue;
                        flashMessage = "Increase ScrollSpeed  to " + newScalarValue.ToString();
                        break;
                    case ApiMessage.ScrollSpeedDecrease:
                        // Can we decrease the ScrollSpeed by the requested amount?
                        // If so do it. Otherwise go as low as settings will allow.
                        newScalarValue = Math.Round(Math.Max(ScrollSpeedSlider.Value - message.Scalar,
                            Properties.Settings.Default.LowerScrollLimit), 2);
                        ScrollSpeedSlider.Value = newScalarValue;
                        flashMessage = "Decrease Scroll Speed to " + newScalarValue.ToString();
                        break;
                    case ApiMessage.ScrollSpeedReset:
                        ScrollSpeedSlider.Value = Properties.Settings.Default.DefaultScrollVelocity;
                        flashMessage = "Reset Scroll Speed to default";
                        break;
                    default:
                        break;
                }
                break;
            case ApiMessageVerb.Spacing:
                switch (message.Parameters)
                {
                    case ApiMessage.SpacingIncrease:
                        // Can we increase the Spacing by the requested amount?
                        // If so do it. Otherwise go as high as settings will allow.
                        newScalarValue = Math.Round(Math.Min(LineSpacingSlider.Value + message.Scalar,
                            Properties.Settings.Default.UpperLineSpacingLimit), 2);
                        LineSpacingSlider.Value = newScalarValue;
                        flashMessage = "Increase Line Spacing to " + newScalarValue.ToString();
                        break;
                    case ApiMessage.SpacingDecrease:
                        // Can we decrease the Spacing by the requested amount?
                        // If so do it. Otherwise go as low as common sense will allow.
                        newScalarValue = Math.Round(Math.Max(LineSpacingSlider.Value - message.Scalar, 1.0), 2);
                        LineSpacingSlider.Value = newScalarValue;
                        flashMessage = "Decrease Line Spacing to " + newScalarValue.ToString();
                        break;
                    case ApiMessage.SpacingReset:
                        LineSpacingSlider.Value = Properties.Settings.Default.DefaultLineSpacing;
                        flashMessage = "Reset Line Spacing to default";
                        break;
                    default:
                        break;
                }
                break;
            case ApiMessageVerb.Margin:
                switch (message.Parameters)
                {
                    case ApiMessage.MarginIncrease:
                        // Can we increase the Margin by the requested amount?
                        // If so do it. Otherwise go as high as settings will allow.
                        newScalarValue = Math.Round(Math.Min(SideMarginSlider.Value + message.Scalar,
                            Properties.Settings.Default.UpperHorizontalMarginLimit), 2);
                        SideMarginSlider.Value = newScalarValue;
                        flashMessage = "Increase Horizontal Margin to " + newScalarValue.ToString();
                        break;
                    case ApiMessage.MarginDecrease:
                        // Can we decrease the Margin by the requested amount?
                        // If so do it. Otherwise go as low as common sense will allow.
                        newScalarValue = Math.Round(Math.Max(SideMarginSlider.Value - message.Scalar, 1.0), 2);
                        SideMarginSlider.Value = newScalarValue;
                        flashMessage = "Decrease Horizontal Margin to " + newScalarValue.ToString();
                        break;
                    case ApiMessage.MarginReset:
                        SideMarginSlider.Value = Properties.Settings.Default.DefaultHorizontalMargin;
                        flashMessage = "Reset Horizontal Margin to default";
                        break;
                    default:
                        break;
                }
                break;
            case ApiMessageVerb.Theme:
                switch (message.Parameters)
                {
                    case ApiMessage.ThemeLight:
                        flashMessage = "Set viewer window to Light Theme";
                        ThemeLightRadioButton.IsChecked = true;
                        _viewerWindow.SetVisualTheme(ViewerWindow.VisualTheme.Light);
                        break;
                    case ApiMessage.ThemeDark:
                        flashMessage = "Set viewer window to Dark Theme";
                        ThemeDarkRadioButton.IsChecked = true;
                        _viewerWindow.SetVisualTheme(ViewerWindow.VisualTheme.Dark);
                        break;
                    case ApiMessage.ThemeMatrix:
                        flashMessage = "Set viewer window to Matrix Theme";
                        ThemeMatrixRadioButton.IsChecked = true;
                        _viewerWindow.SetVisualTheme(ViewerWindow.VisualTheme.Matrix);
                        break;
                    default:
                        break;
                }
                break;
            case ApiMessageVerb.Bug:
                switch (message.Parameters)
                {
                    case ApiMessage.BugNone:
                        flashMessage = "Set progress bug display to None";
                        ProgressBugNoneRadioButton.IsChecked = true;
                        _viewerWindow.SetProgressBug(ViewerWindow.ProgressBug.None);
                        break;
                    case ApiMessage.BugSide:
                        flashMessage = "Set progress bug display to Side";
                        ProgressBugSideRadioButton.IsChecked = true;
                        _viewerWindow.SetProgressBug(ViewerWindow.ProgressBug.RightSide);
                        break;
                    case ApiMessage.BugBottom:
                        flashMessage = "Set progress bug display to Bottom";
                        ProgressBugBottomRadioButton.IsChecked = true;
                        _viewerWindow.SetProgressBug(ViewerWindow.ProgressBug.Bottom);
                        break;
                    default:
                        break;
                }
                break;
            default:
                break;
        }

        if (!string.IsNullOrEmpty(flashMessage))
        {
            FlashPlaybackMessage(flashMessage);
        }
    }

    /// <summary>
    /// Temporarily displays a message on the UI before fading it out.
    /// Used to indicate when a background action (like an API command) takes place
    /// </summary>
    /// <param name="message">The message to be shown</param>
    private void FlashPlaybackMessage(string message)
    {
        LowerRightCornerStatusMessages.Text = message;

        // Animate the PlaybackTabStatusMessages TextBlock opacity
        SineEase easing = new()
        {
            EasingMode = EasingMode.EaseOut
        };
        // NOTE: Want to go from 1.0 to 0.0 but start at 4.0 so the first 3/4 isn't effectively animating (yet)
        DoubleAnimation fadeOutAnimation = new(4.0, 0.0, TimeSpan.FromMilliseconds(2000))
        {
            EasingFunction = easing
        };

        LowerRightCornerStatusMessages.BeginAnimation(UIElement.OpacityProperty, fadeOutAnimation, HandoffBehavior.SnapshotAndReplace);
    }

    private void ThemeRadioButton_Click(object sender, RoutedEventArgs e)
    {
        if (ThemeLightRadioButton.IsChecked == true)
        {
            _viewerWindow.SetVisualTheme(ViewerWindow.VisualTheme.Light);
        }
        else if (ThemeDarkRadioButton.IsChecked == true)
        {
            _viewerWindow.SetVisualTheme(ViewerWindow.VisualTheme.Dark);
        }
        else if (ThemeMatrixRadioButton.IsChecked == true)
        {
            _viewerWindow.SetVisualTheme(ViewerWindow.VisualTheme.Matrix);
        }
        // Otherwise: Shouldn't happen, but do nothing if it does.
    }

    private void ProgressBugRadioButton_Click(object sender, RoutedEventArgs e)
    {
        if (ProgressBugNoneRadioButton.IsChecked == true)
        {
            _viewerWindow.SetProgressBug(ViewerWindow.ProgressBug.None);
        }
        else if (ProgressBugSideRadioButton.IsChecked == true)
        {
            _viewerWindow.SetProgressBug(ViewerWindow.ProgressBug.RightSide);
        }
        else if (ProgressBugBottomRadioButton.IsChecked == true)
        {
            _viewerWindow.SetProgressBug(ViewerWindow.ProgressBug.Bottom);
        }
        // Otherwise: Shouldn't happen, but do nothing if it does.
    }

    /// <summary>
    /// Attempts to establish a connection to the ATEM video switcher and read a set of available inputs
    /// </summary>
    /// <param name="atemIpAddress">The presumed IP address of the ATEM</param>
    private void ConnectToAtem(string atemIpAddress)
    {
        _atemSwitcher = new AtemSwitcher(atemIpAddress: atemIpAddress);
        if (!_loading)
        {
            if (_atemSwitcher is not null && _atemSwitcher.IsReady)
            {
                AtemApiStatusLed.SelectedColour = LightEmittingDiode.ColourOptions.Green;
                AtemInputsComboBox.IsEnabled = true;
                // Fill the inputs combo box
                AtemInputsComboBox.ItemsSource = _atemSwitcher.GetInputShortNames(AtemSwitcherPortType.External);
                ComboBoxItem item = (ComboBoxItem)AtemInputsComboBox.FindName(Properties.Settings.Default.ViewerAtemInput);
                if (item is not null)
                {
                    AtemInputsComboBox.SelectedItem = item;
                }
                else
                {
                    AtemInputsComboBox.SelectedIndex = 0;
                }
                Properties.Settings.Default.AtemIpAddress = atemIpAddress;
            }
            else
            {
                AtemApiStatusLed.SelectedColour = LightEmittingDiode.ColourOptions.Red;
                AtemInputsComboBox.IsEnabled = false;
            }
        }
    }

    private void AtemInputsComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!_loading && AtemInputsComboBox.SelectedItem is not null)
        {
            Properties.Settings.Default.ViewerAtemInput = AtemInputsComboBox.SelectedItem.ToString();
        }
    }

    private void AtemConnectButton_Click(object sender, RoutedEventArgs e)
    {
        using (new WaitCursor())
        {
            ConnectToAtem(AtemIpAddressTextBox.Text.Trim());
        }
    }


}
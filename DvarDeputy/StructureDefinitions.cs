namespace Mooseware.DvarDeputy;

/// <summary>
/// A single text fragment (usually, but not necessarily a word) that is displayed in the teleprompter. 
/// </summary>
/// <param name="content">The textual content of the fragment</param>
/// <param name="start">The starting index for the fragment in the main ReadOnlySpan(char)</param>
/// <param name="count">The count of chars in the fragment (used with Start to Slice from the span)</param>
/// <param name="bold">True if this fragment is to be formatted as bold text when displayed</param>
/// <param name="italic">True if this fragment is to be formatted as italic text when displayed</param>
/// <param name="underline">True if this fragment is to be formatted with underline when displayed</param>
/// <param name="suppressTrailingSpace">False if the fragment is to be followed by a trailing space and
///  true if the fragment is to have the trailing space suppressed (e.g. end of line, partial word bold/italic &c.
///  (default=false)</param>
///  <param name="isLineBreak">Whether the fragment ends in a line break</param>
///  <param name="isEmpty">Whether the fragment is empty (other than, perhaps, a line break)</param>
///  <param name="hasHebrew">Whether or not the text fragment contains Hebrew characters</param>
///  <param name="numberingId">Which numbering scheme applies to the fragment. 
///  When greater than 0, indicates the fragment is the start of bullets or numbering</param>
///  <param name="numberingLevel">Which numbering level (depth) applies to the fragment. 
///  When greater than 0, indicates the fragment is the start of bullets or numbering</param>
internal readonly struct TextFragment(
    string content, 
    bool bold = false, 
    bool italic = false, 
    bool underline = false, 
    bool suppressTrailingSpace = false, 
    bool isLineBreak = false,
    bool isEmpty = false,
    bool hasHebrew = false,
    int numberingId = 0,
    int numberingLevel = 0)
{
    /// <summary>
    /// The string content of the fragment
    /// </summary>
    public string Content { get; } = content;

    /// <summary>
    /// True if this fragment is to be formatted as bold text when displayed, false otherwise
    /// </summary>
    public bool Bold { get; } = bold;

    /// <summary>
    /// True if this fragment is to be formatted as italic text when displayed, false otherwise
    /// </summary>
    public bool Italic { get; } = italic;

    /// <summary>
    /// True if this fragment is to be formatted with underline when displayed, false otherwise
    /// </summary>
    public bool Underline { get; } = underline;

    /// <summary>
    /// False if the fragment is to be followed by a trailing space and true if the fragment is 
    /// to have the trailing space suppressed (e.g. end of line, partial word bold/italic &c. (default=false)
    /// </summary>
    public bool SuppressTrailingSpace { get; } = suppressTrailingSpace;

    /// <summary>
    /// Whether the fragment ends in a line break
    /// </summary>
    public bool IsLineBreak { get; } = isLineBreak;

    /// <summary>
    /// Whether the fragment is empty (other than, perhaps, a line break)
    /// </summary>
    public bool IsEmpty { get; } = isEmpty;

    /// <summary>
    /// Whether or not the text fragment contains Hebrew characters. May result in lines being rendered RTL instead of LTR.
    /// </summary>
    public bool HasHebrew { get; } = hasHebrew;

    /// <summary>
    /// Which numbering scheme applies to the fragment. 
    ///  When greater than 0, indicates the fragment is the start of bullets or numbering
    /// </summary>
    public int NumberingId { get; } = numberingId;

    /// <summary>
    /// Which numbering level (depth) applies to the fragment. 
    ///  When greater than 0, indicates the fragment is the start of bullets or numbering
    /// </summary>
    public int NumberingLevel { get; } = numberingLevel;
}

/// <summary>
/// A single line of text to be displayed in a teleprompter screen based on a pointer to a
/// List<TextFragment> containing the text fragments (usually words) that form each line.
/// A List of these LineOfText structs represents the whole content to be displayed.
/// </summary>
/// <param name="start">Index of the first TextFragment at the beginning of the line of text</param>
/// <param name="count">Number of TextFragement structs which make up a line of text</param>
/// <param name="isRightToLeft">Whether the line is to be painted RTL (e.g. it's Hebrew)</param>
/// <param name="margin">The amount of indenting at the start of the line (usually left margin)</param>
/// <param name="tabStop">The amount of space to render between a bullet/number and the start of the bulleted/numbered text</param>
internal readonly struct LineOfText(int start = 0, int count = 1, bool isRightToLeft = false, double margin = 0.0, double tabStop = 0.0)
{
    /// <summary>
    /// Index of the first TextFragment at the beginning of the line of text
    /// </summary>
    public int Start { get; } = start;
    /// <summary>
    /// Number of TextFragement structs which make up a line of text
    /// </summary>
    public int Count { get; } = count;
    /// <summary>
    /// Whether the line is to be rendered in Right to Left direction (c.f. Left to Right)
    /// </summary>
    public bool IsRightToLeft { get; } = isRightToLeft;
    /// <summary>
    /// The amount of indenting at the start of the line (usually left margin)
    /// </summary>
    public double Margin { get; } = margin;
    /// <summary>
    /// The amount of space to render between a bullet/number and the start of the bulleted/numbered text
    /// </summary>
    public double TabStop { get; } = tabStop;
}

/// <summary>
/// A structure that contains character formatting flags for the types of formatting
/// supported by the application for prompted content text
/// </summary>
internal record CharacterFormattingProperties
{
    /// <summary>
    /// Is the character string rendered in Bold type weight
    /// </summary>
    public bool Bold { get; set; } = false;
    /// <summary>
    /// Is the character string rendered in Italic font style
    /// </summary>
    public bool Italic { get;set; } = false;
    /// <summary>
    /// Is the character string rendered with underline text decoration
    /// </summary>
    public bool Underline { get; set; } = false;
    public CharacterFormattingProperties()
    {
    }
}

/// <summary>
/// Supported formats for Word document content bullet and numbering schemes
/// </summary>
internal enum WordNumberingFormat
{
    /// <summary>
    /// Scheme to be determined
    /// </summary>
    Unknown,
    /// <summary>
    /// Bullets
    /// </summary>
    Bullet,
    /// <summary>
    /// Decimal numbers (1,2,3,...)
    /// </summary>
    Decimal,
    /// <summary>
    /// Uppercase letters (A,B,C,...)
    /// </summary>
    UpperLetter,
    /// <summary>
    /// Lowercase letters (a,b,c,...)
    /// </summary>
    LowerLetter,
    /// <summary>
    /// Uppercase Roman numerals (I,II,III,...)
    /// </summary>
    UpperRoman,
    /// <summary>
    /// Lowercase Roman numberals (i,ii,iii,...)
    /// </summary>
    LowerRoman
}

/// <summary>
/// The details of a Word document Numbering scheme (in particular an Abstract Numbering 
/// Scheme) as identified by the Numbering ID of a Numbering Instance and the Level ID
/// </summary>
internal record WordNumberingScheme
{
    /// <summary>
    /// The Numbering Instance NumberID value. Used as part of the key of this structure
    /// </summary>
    internal int NumberId { get; set; } = 0;
    /// <summary>
    /// The Level of the Abstract Numbering scheme item representing the depth of the bullet/number
    /// </summary>
    internal int LevelIndex { get; set; } = 0;
    /// <summary>
    /// The AbstractNum.AbstractNumberId of the scheme
    /// </summary>
    internal int AbstractNumberingId { get; set; } = 0;
    /// <summary>
    /// The number from which to start a numbering paragraphs in this scheme
    /// </summary>
    internal int StartingNumber { get; set; } = 0;
    /// <summary>
    /// A WordNumberingFormat enum value based on the NumberingFormat of the abstract numbering scheme
    /// </summary>
    internal WordNumberingFormat Format { get; set; } = WordNumberingFormat.Unknown;
    /// <summary>
    /// The value of the Abstract numbering scheme's NumberingFormat, which gives the structure of the numbering.
    /// Levels in numeric templates show the depth like so: "%1", "%2", etc. Other characters in the template
    /// are to be treated as literal text symbols in the numbering representation (typically punctuation: '.',')' &c.)
    /// </summary>
    internal string Template { get; set; } = string.Empty;
    /// <summary>
    /// The symbol to use for bullets in a bulleted numbering scheme
    /// </summary>
    internal string Symbol { get; set; } = string.Empty;
    /// <summary>
    /// Whether or not the Format is one of the numbering based formats (c.f. bullets or unknown)
    /// </summary>
    internal bool IsNumbering
    {
        get
        {
            return !(Format == WordNumberingFormat.Bullet || Format == WordNumberingFormat.Unknown);
        }
    }
}
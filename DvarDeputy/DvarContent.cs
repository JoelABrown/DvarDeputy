using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Media;

namespace Mooseware.DvarDeputy;

/// <summary>
/// The content of a document to be prompted with properties and methods that facilitate prompting operations.
/// </summary>
internal class DvarContent
{
    /////// <summary>
    /////// The full path and file containing the loaded content.
    /////// </summary>
    ////private string _contentFilespec = string.Empty;

    /// <summary>
    /// A list of TextFragments which comprise the content to be prompted
    /// </summary>
    internal List<TextFragment> Fragments { get; private set; } = [];

    /// <summary>
    /// A list of LinesOfText which comprise the content to be prompted as constituted into lines based on given visual properties
    /// </summary>
    internal List<LineOfText> LinesOfText { get; private set; } = [];

    /// <summary>
    /// Word document numbering schemes keyed by the Numbering Instance's NumberID and
    /// the Abstract Numbering's Level Index. Provides details of the numbering (or bullet)
    /// paragraph's properties.
    /// </summary>
    private Dictionary<(int, int), WordNumberingScheme> NumberingSchemes { get; set; } = [];

    /// <summary>
    /// The full path and filespec used to source the raw content
    /// </summary>
    private string _rawContentFilespec = string.Empty;

    /// <summary>
    /// The full path and filespec of the currently loaded content file
    /// </summary>
    public string RawContentFilespec { get => _rawContentFilespec; }

    /// <summary>
    /// The factor by which hanging indents are multiplied to account for the space needed for bullets and numbering
    /// </summary>
    internal static double HangingIndentFactor { get => 1.1; }

    internal DvarContent()
    {

    }

    /// <summary>
    /// Loads content from a given file of one of the supported file formats
    /// </summary>
    /// <param name="contentFilespec">The full path and file specification of the file to be loaded</param>
    internal void LoadContentFromFile(string contentFilespec)
    {
        _rawContentFilespec = string.Empty; // Until we find out that the file is there.

        // Reset existing content
        LinesOfText.Clear();
        Fragments.Clear();

        // Make sure the file exists
        if (!File.Exists(contentFilespec))
        {
            return;
        }

        // Determine what kind of file we have and handle it accordingly
        string extension = System.IO.Path.GetExtension(contentFilespec);
        switch (extension.ToLower())
        {
            case ".txt":
                ParseContentFromTextFile(contentFilespec);
                break;
            case ".docx":
                ParseContentFromWordDocument(contentFilespec);
                break;
            default:
                break;
        }
    }

    /// <summary>
    /// Parses text file content into a list of TextFragments (typically words) with given styling options indicated
    /// </summary>
    /// <param name="filespec">The full path and file of the content to be loaded and parsed into fragments</param>
    private void ParseContentFromTextFile(string filespec)
    {
        // Sanity check
        if (!File.Exists(filespec)) return;
        _rawContentFilespec = filespec;

        // Read the file and retrieve all of its contents
        // Work with a span of char to speed things up a little bit at least
        var rawContent = File.ReadAllText(filespec, System.Text.Encoding.UTF8).AsSpan();

        // Guestimate the required capacity for the list of text fragments based on
        // a presumed average word length. This will be trimmed later.
        int estimatedWords = Math.Max((int)(rawContent.Length / 5), 10);
        Fragments = new List<TextFragment>(estimatedWords);

        // Open up the first fragment
        bool isBold = false;
        bool isItalic = false;
        bool isUnderline = false;
        bool suppressTrailingSpace = false;
        bool isLineBreak = false;
        int startIdx = 0;
        int length = 0;
        bool fragmentComplete = false;
        bool hasHebrewChacters = false;
        // Look at each character in turn and break the raw content into fragments
        for (int i = 0; i < rawContent.Length; i++)
        {
            char thisChar = rawContent[i];
            hasHebrewChacters |= IsAHebrewCharacter(thisChar);
            switch (thisChar)
            {
                case ' ':
                    fragmentComplete = true;
                    break;

                // TODO: Determine whether to trap '\t' and handle bullets or at least convert it to a space.

                case '\r':
                    fragmentComplete = true;
                    isLineBreak = true;
                    suppressTrailingSpace = true;
                    // Check to see if this \r is followed immediately by a \n.
                    // If so, suppress this char and let the \n do the job.
                    if (i < rawContent.Length - 1)
                    {
                        if (rawContent[i + 1] == '\n')
                        {
                            // Skip ahead by one character
                            i++;
                        }
                    }
                    break;

                case '\n':
                    fragmentComplete = true;
                    isLineBreak = true;
                    suppressTrailingSpace = true;
                    break;

                default:
                    // Include this in the fragment
                    length++;
                    break;
            }

            if (fragmentComplete)
            {
                // If the fragment is a line break with empty or a blank content then flag that
                // since empty fragments aren't meant to be run together in long runs...
                bool isEmpty = (isLineBreak && rawContent.Slice(startIdx, length).ToString().Trim().Length == 0);

                // Skip empty (but complete) fragments unless they are line breaks
                if (length > 0 || isLineBreak == true)
                {
                    // Record the fragment
                    TextFragment fragment = new(
                        rawContent.Slice(startIdx, length).ToString(),
                        isBold,
                        isItalic,
                        isUnderline,
                        suppressTrailingSpace,
                        isLineBreak,
                        isEmpty,
                        hasHebrewChacters);

                    // If the fragment has Hebrew characters check to see if the last character
                    // is actually punctuation. HE text pastes into text files with the punctuation
                    // at the wrong end of the word because of the vaguaries of mixed LTR/RTL text handling.
                    if (hasHebrewChacters)
                    {
                        string lastChar = fragment.Content.Trim()[^1].ToString();
                        // Is the last character one of the expected / handled punctuation marks?
                        if ((",.!?").Contains(lastChar))
                        {
                            // Move the punctuation to the beginning of the string.
                            string punctuation = lastChar;
                            string word = fragment.Content.Trim()[..^1];
                            string correctedContent = punctuation + word;
                            // Reconstitute the fragment with the corrected content.
                            fragment = new(
                                correctedContent,
                                isBold,
                                isItalic,
                                isUnderline,
                                suppressTrailingSpace,
                                isLineBreak,
                                isEmpty,
                                hasHebrewChacters);
                        }
                    }
                    Fragments.Add(fragment);
                }
                // Reset for the next fragment
                isBold = false;
                isItalic = false;
                isUnderline = false;
                suppressTrailingSpace = false;
                isLineBreak = false;
                hasHebrewChacters = false;
                startIdx = i + 1;
                length = 0;
                fragmentComplete = false;
            }
        }

        // Record the last fragment
        suppressTrailingSpace = true;
        TextFragment lastFragment = new(
            rawContent.Slice(startIdx, length).ToString(),
            isBold,
            isItalic,
            isUnderline,
            suppressTrailingSpace,
            true,   // Is a line break (by definition)
            false,  // Is not empty (or don't care if it is)
            hasHebrewChacters
            );
        Fragments.Add(lastFragment);

        Fragments.TrimExcess();
    }

    /// <summary>
    /// Parses Word (.docx) file content into a list of TextFragments (typically words) with given styling options indicated
    /// </summary>
    /// <param name="filespec">The full path and file of the content to be loaded and parsed into fragments</param>
    private void ParseContentFromWordDocument(string filespec)
    {
        // -----------------------------------------------------------------------------
        // This is the structure of a Word document (at least the part we care about)
        //
        // {Body}      (Under the MainDocumentPart.Document...)
        // +-{Paragraph}
        //   +-{Paragraph Properties} (opt)
        //   | +-{Numbering Properties}
        //   |   +-{Numbering Level Reference}
        //   |   +-{Numbering ID}
        //   +-{Run}
        //     +-{Run Properties} (opt)
        //     | +-{Bold}
        //     | +-{Italic}
        //     | +-{Underline}
        //     +-{Text}
        //
        // NOTES:
        // - Paragraphs always end in a line break (obviously)
        // - Bullets, numbering and indents are controlled by paragraph properites
        // - Runs all have the same character formatting
        // - Runs can contain multiple fragmens and a fragment can span multiple runs
        //   Therefore, runs need to be concatenated together as long as they have
        //   consistent formatting. Once the formatting changes or the paragraph ends,
        //   then fragments need to be extracted from that whole concatenated string.
        //
        // FOR BULLETS AND NUMBERING:
        // Numbering formats are stored separately in the .docx file and are referenced
        // in the body content via Paragraph Properties. Each numbering style has an
        // instance reference which itself cross-references to an abstract style 
        // definition, where the actual number/bullet styling details are stored.
        // This is the relevant part of the structure for bullet and numbering styles:
        //
        // {Numbering}     (Under MainDocumentPart.NumberingDefinitionsPart)
        // +-{Numbering Instance {Number ID}} -----> From Doc Paragraph Properties.
        // | +-{Abstract Num {Abstract Num ID}}----+
        // |                                       | Points to this
        // +-{Abstract Num {Abstract Num ID}} <----+
        //   +-{Level {Level Index}} --------------> From Doc Paragraph Properties.
        //     +-{Level Start}
        //     +-{Numbering Format} 
        //     +-{Level Text}
        //
        // NOTES:
        // - Numbering Format might be "bullet" or a numbering scheme like "decimal" or "lowerLetter" or "lowerRoman"
        // - Level Text is the bullet character or a number format like "%1."
        //   (in "%1." the "%1" indicates the depth of the nested counter)
        // - The Level node under Abstract Num can have a Run Properties node if the Numbering Format is bullet.
        //   This gives information about the font to use for the Level Text (i.e. the bullet character)
        //   => See if this can be ignored for now!
        // - The Level node under Abstract Num also has Paragraph Properties with margin and hanging properties
        //   but these should probably be ignored for purposes of prompting text.
        //
        // -----------------------------------------------------------------------------

        // Sanity check
        if (!File.Exists(filespec)) return;
        _rawContentFilespec = filespec;

        // Open the document in read-only mode
        using WordprocessingDocument wordDoc = WordprocessingDocument.Open(filespec, false);

        // Bail out if the document open failed for whatever reason
        if (wordDoc == null
          || wordDoc.MainDocumentPart == null
          || wordDoc.MainDocumentPart.Document == null
          || wordDoc.MainDocumentPart.Document.Body == null) return;

        // Be ready to track nested numbering.
        int[] numberingSequences = [0, 0, 0];    // Max of three levels supported for prompting!

        // Open the body of the word document and look for (and process) each paragraph and its contents
        var body = wordDoc.MainDocumentPart.Document.Body;
        foreach (var bodyChild in body)
        {
            if (bodyChild.GetType() == typeof(DocumentFormat.OpenXml.Wordprocessing.Paragraph))
            {
                var paragraph = (DocumentFormat.OpenXml.Wordprocessing.Paragraph)bodyChild;
                if (paragraph == null) continue;
                CharacterFormattingProperties priorRunFormatting = new();
                string collectiveRunContent = string.Empty;
                bool thisParagraphHasNumbering = false;

                // Look at the children of the paragraph
                foreach (var paragraphChild in paragraph.ChildElements)
                {
                    // Determine which kind of paragraph child element we have
                    if (paragraphChild.GetType() == typeof(DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties))
                    {
                        int numberingLevel = -1;
                        int numberingId = -1;
                        foreach (var paraPropsChild in paragraphChild.ChildElements)
                        {
                            if (paraPropsChild.GetType() == typeof(DocumentFormat.OpenXml.Wordprocessing.NumberingProperties))
                            {
                                var numberingProperties = (DocumentFormat.OpenXml.Wordprocessing.NumberingProperties)paraPropsChild;
                                foreach (var numPropsChild in numberingProperties.ChildElements)
                                {
                                    // Note the paragraph properties, particularly if they have to do with bullets or numbering.
                                    if (numPropsChild.GetType() == typeof(DocumentFormat.OpenXml.Wordprocessing.NumberingLevelReference))
                                    {
                                        var numLvlRef = (DocumentFormat.OpenXml.Wordprocessing.NumberingLevelReference)numPropsChild;
                                        if (numLvlRef is not null && numLvlRef.Val is not null && numLvlRef.Val.HasValue)
                                        {
                                            thisParagraphHasNumbering = true;
                                            numberingLevel = (int)numLvlRef.Val.Value;
                                        }
                                    }
                                    else if (numPropsChild.GetType() == typeof(DocumentFormat.OpenXml.Wordprocessing.NumberingId))
                                    {
                                        var numIdRef = (DocumentFormat.OpenXml.Wordprocessing.NumberingId)numPropsChild;
                                        if (numIdRef is not null && numIdRef.Val is not null && numIdRef.Val.HasValue)
                                        {
                                            thisParagraphHasNumbering = true;
                                            numberingId = (int)numIdRef.Val.Value;
                                        }
                                    }
                                }
                            }
                        }
                        // If we found numbering in the paragraph properties, we need to 
                        // make sure that we make a note of the numbering settings being used
                        if (thisParagraphHasNumbering && numberingId > 0 && numberingLevel >= 0)
                        {
                            // Get the bullet or numbering scheme
                            LookUpWordNumberingScheme(wordDoc, numberingId, numberingLevel);
                            var numScheme = GetNumberingScheme(numberingId, numberingLevel);
                            string bulletNumberContent = string.Empty; 
                            if (numScheme.IsNumbering)
                            {
                                // If it's numbering, (c.f. bullets), increment/reset the sequence counters based on the level.
                                switch (numberingLevel)
                                {
                                    case 0:
                                        numberingSequences[0]++;
                                        numberingSequences[1] = 0;
                                        numberingSequences[2] = 0;
                                        break;
                                    case 1:
                                        numberingSequences[1]++;
                                        numberingSequences[2] = 0;
                                        break;
                                    case 2:
                                        numberingSequences[2]++;
                                        break;
                                    default:
                                        // Not supported.
                                        break;
                                }
                                bulletNumberContent = NumberingValue(numScheme, numberingSequences);
                            }
                            else
                            {
                                bulletNumberContent = numScheme.Symbol;
                            }
                            // Create a start of numbering/bulletting fragment
                            TextFragment numberingFragment = new(content: bulletNumberContent,
                                                                 suppressTrailingSpace: true,
                                                                 numberingId: numberingId,
                                                                 numberingLevel: numberingLevel);
                            Fragments.Add(numberingFragment);
                        }
                    }
                    else if (paragraphChild.GetType() == typeof(DocumentFormat.OpenXml.Wordprocessing.Run))
                    {
                        // Process a Wordprocessing.Run node.
                        string thisRunContent = string.Empty;
                        CharacterFormattingProperties thisRunFormatting = new();
                        foreach (var runChild in paragraphChild.ChildElements)
                        {
                            if (runChild.GetType() == typeof(DocumentFormat.OpenXml.Wordprocessing.RunProperties))
                            {
                                var runProps = (DocumentFormat.OpenXml.Wordprocessing.RunProperties)runChild;
                                thisRunFormatting.Bold = (runProps.Bold is not null);
                                thisRunFormatting.Italic = (runProps.Italic is not null);
                                thisRunFormatting.Underline = (runProps.Underline is not null);
                            }
                            else if (runChild.GetType() == typeof(DocumentFormat.OpenXml.Wordprocessing.Text))
                            {
                                var runText = (DocumentFormat.OpenXml.Wordprocessing.Text)runChild;
                                thisRunContent = runText.Text;
                            }
                        }
                        if (thisRunFormatting == priorRunFormatting)
                        {
                            // We can add this to the "pile" and keep going.
                            collectiveRunContent += thisRunContent;
                        }
                        else
                        {
                            // This represents a break in the run formatting.
                            // The current collected run needs to be parsed for fragments
                            if (collectiveRunContent.Length > 0)
                            {
                                ParseCollectedDocxRun(collectiveRunContent, priorRunFormatting);
                            }

                            // Set up for a new collective run starting with this 
                            // latest run with differentiated formatting
                            collectiveRunContent = thisRunContent;
                            // Note the new formatting for the next run
                            priorRunFormatting = thisRunFormatting;
                        }
                    }
                }
                // Parse any remaining runs that have yet to be recorded.
                if (collectiveRunContent.Length > 0)
                { 
                    ParseCollectedDocxRun(collectiveRunContent, priorRunFormatting); 
                }
                // At the end of a paragraph, create an empty end of line fragment
                TextFragment endOfParagraphfragment = new(string.Empty,false,false,false,true,true,true,false);
                Fragments.Add(endOfParagraphfragment);

                if (!thisParagraphHasNumbering)
                {
                    // Reset the numbering scheme counters
                    numberingSequences = [0, 0, 0];
                }
            }
        }
    }

    /// <summary>
    /// Takes a single line of text collected from parsing runs from a Word .docx file
    /// and parses it into the appropriate number of Fragments with consistent character
    /// formatting. Note that end of line fragments won't come up because that's a
    /// paragraph concern in .docx files, not a run concern.
    /// </summary>
    /// <param name="content">The collected string to be parsed</param>
    /// <param name="formatting">The character formatting to be applied</param>
    private void ParseCollectedDocxRun(string content, CharacterFormattingProperties formatting)
    {
        var rawContent = content.AsSpan();
        int startIdx = 0;
        int length = 0;
        bool fragmentComplete = false;
        bool hasHebrewChacters = false;
        bool suppressTrailingSpace = false;
        // Look at each character in turn and break the raw content into fragments
        for (int i = 0; i < rawContent.Length; i++)
        {
            char thisChar = rawContent[i];
            hasHebrewChacters |= IsAHebrewCharacter(thisChar);
            switch (thisChar)
            {
                case ' ':
                    // NOTE: Sometimes a fragment can start with a leading space
                    //       because the previous fragment had different character
                    //       formatting properties. Don't suppress leading spaces.
                    if (i == startIdx)
                    {
                        length++;
                        suppressTrailingSpace = true;   // But don't add a space to the space!
                    }
                    fragmentComplete = true;
                    break;

                // TODO: Determine whether to trap '\t' inside of Word document content.

                default:
                    // Include this in the fragment
                    length++;
                    break;
            }

            if (i == rawContent.Length - 1)
            {
                fragmentComplete = true;
                suppressTrailingSpace = (thisChar != ' ');
            }

            if (fragmentComplete)
            {
                // If the fragment is a line break with empty or a blank content then flag that
                // since empty fragments aren't meant to be run together in long runs...
                bool isEmpty = (length == 0);

                // Skip empty (but complete) fragments unless they are line breaks
                if (length > 0)
                {
                    // Record the fragment
                    TextFragment fragment = new(
                        rawContent.Slice(startIdx, length).ToString(),
                        formatting.Bold,
                        formatting.Italic,
                        formatting.Underline,
                        suppressTrailingSpace,
                        false,
                        isEmpty,
                        hasHebrewChacters);

                    // If the fragment has Hebrew characters check to see if the last character
                    // is actually punctuation. HE text pastes into text files with the punctuation
                    // at the wrong end of the word because of the vaguaries of mixed LTR/RTL text handling.
                    if (hasHebrewChacters)
                    {
                        string lastChar = fragment.Content.Trim()[^1].ToString();
                        // Is the last character one of the expected / handled punctuation marks?
                        if ((",.!?").Contains(lastChar))
                        {
                            // Move the punctuation to the beginning of the string.
                            string punctuation = lastChar;
                            string word = fragment.Content.Trim()[..^1];
                            string correctedContent = punctuation + word;
                            // Reconstitute the fragment with the corrected content.
                            fragment = new(
                                correctedContent,
                                formatting.Bold,
                                formatting.Italic,
                                formatting.Underline,
                                suppressTrailingSpace,
                                false,
                                isEmpty,
                                hasHebrewChacters);
                        }
                    }
                    Fragments.Add(fragment);
                }
                // Reset for the next fragment
                suppressTrailingSpace = false;
                hasHebrewChacters = false;
                startIdx = i + 1;
                length = 0;
                fragmentComplete = false;
            }
        }
    }

    /// <summary>
    /// Extracts a Word Numbering Scheme from a given Word document based on the numbering instance ID
    /// and the abstract numbering scheme level identifier
    /// </summary>
    /// <param name="wordDoc">The open Word document reference to be searched</param>
    /// <param name="id">The Numbering Instance's NumberID value</param>
    /// <param name="level">The Abstract numbering scheme's Level ID</param>
    private void LookUpWordNumberingScheme(WordprocessingDocument wordDoc, int id, int level)
    {
        // Gottem or needem?
        if (NumberingSchemes.ContainsKey((id,level)))
        {
            // Gottem
            return;
        }

        // Otherwise, needem. Do the lookup...
        if (wordDoc is null
         || wordDoc.MainDocumentPart is null
         || wordDoc.MainDocumentPart.NumberingDefinitionsPart is null
         || wordDoc.MainDocumentPart.NumberingDefinitionsPart.Numbering is null
         || !wordDoc.MainDocumentPart.NumberingDefinitionsPart.Numbering.HasChildren) { return; }

        int abstractId = -1;
        string numberingFormat = string.Empty;
        int start = -1;
        string levelText = string.Empty;
        string bulletChar = "\u2022";

        // Find the numbering child element that is a NumberingInstance type
        foreach (var numberingChild in wordDoc.MainDocumentPart.NumberingDefinitionsPart.Numbering.ChildElements)
        {
            if (numberingChild.GetType() == typeof(DocumentFormat.OpenXml.Wordprocessing.NumberingInstance))
            {
                var numInstance = (DocumentFormat.OpenXml.Wordprocessing.NumberingInstance)numberingChild;
                if (numInstance.NumberID is not null)
                {
                    int numId = (int)numInstance.NumberID;
                    if (numId == id)
                    {
                        if (numInstance.AbstractNumId is not null)
                        {
                            abstractId = numInstance.AbstractNumId.Val ?? -1;
                        }
                        break;
                    }
                }
            }
        }

        // For the instance, get the abstract type
        foreach (var numberingChild in wordDoc.MainDocumentPart.NumberingDefinitionsPart.Numbering.ChildElements)
        {
            if (numberingChild.GetType() == typeof(DocumentFormat.OpenXml.Wordprocessing.AbstractNum))
            {
                var numInstance = (DocumentFormat.OpenXml.Wordprocessing.AbstractNum)numberingChild;
                if (numInstance.AbstractNumberId is not null)
                {
                    int abstractNumId = (int)numInstance.AbstractNumberId;
                    if (abstractNumId == abstractId)
                    {
                        if (numInstance.HasChildren)
                        {
                            foreach (var item in numInstance.ChildElements)
                            {
                                if (item.GetType() == typeof(DocumentFormat.OpenXml.Wordprocessing.Level))
                                {
                                    var levelItem = (DocumentFormat.OpenXml.Wordprocessing.Level)item;
                                    if (levelItem is not null
                                        && levelItem.LevelIndex is not null
                                        && levelItem.LevelIndex == level)
                                    {
                                        if (levelItem.LevelText is not null && levelItem.LevelText.Val is not null)
                                        {
                                            // TODO: Figure out what to do about applying font from the Word docx to bullets instead
                                            //       of just using the default constant "\u2022"
                                            // bulletChar = levelItem.LevelText.Val.ToString() ?? string.Empty;
                                        }
                                        if (levelItem.NumberingFormat is not null && levelItem.NumberingFormat.Val is not null)
                                        {
                                            numberingFormat = levelItem.NumberingFormat.Val.ToString() ?? string.Empty;
                                        }
                                        if (levelItem.StartNumberingValue is not null && levelItem.StartNumberingValue.Val is not null)
                                        {
                                            start = levelItem.StartNumberingValue.Val;
                                        }
                                        if (levelItem.LevelText is not null && levelItem.LevelText.Val is not null)
                                        {
                                            levelText = levelItem.LevelText.Val.ToString() ?? string.Empty;
                                        }
                                    }
                                }
                            }
                        }
                        break;
                    }
                }
            }
        }

        // Record the WordNumberingScheme
        WordNumberingScheme newScheme = new()
        {
            NumberId = id,
            LevelIndex = level,
            AbstractNumberingId = abstractId,
            StartingNumber = start,
            Format = numberingFormat.ToLower() switch
            {
                "bullet" => WordNumberingFormat.Bullet,
                "decimal" => WordNumberingFormat.Decimal,
                "upperletter" => WordNumberingFormat.UpperLetter,
                "lowerletter" => WordNumberingFormat.LowerLetter,
                "upperroman" => WordNumberingFormat.UpperRoman,
                "lowerroman" => WordNumberingFormat.LowerRoman,
                _ => WordNumberingFormat.Unknown,
            },
            Template = levelText,
            Symbol = bulletChar
        };

        NumberingSchemes.Add((id,level),newScheme);
    }

    internal WordNumberingScheme GetNumberingScheme(int numberingId, int numberingLevel)
    {
        // Gottem or needem?
        if (NumberingSchemes.ContainsKey((numberingId, numberingLevel)))
        {
            // Gottem
            return NumberingSchemes[(numberingId, numberingLevel)];
        }
        else
        {   // This is a fallback in case something went wrong parsing the Word document.
            // Make up a default scheme to use
            WordNumberingScheme defaultScheme = new()
            {
                NumberId = numberingId,
                LevelIndex = numberingLevel,
                AbstractNumberingId = 0,
                StartingNumber = 1,
                Format = WordNumberingFormat.Bullet,
                Template = "\u2022",
                Symbol = "\u2022"
            };
            return defaultScheme;
        }
    }

#if DEBUG
    private static void LogLayoutLineAction(string logText, bool cleanFile = false)
    {
        string filespec = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "dvardeputylayoutline.log");
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
    /// Determines the content of lines based on current content and the provided display parameters
    /// </summary>
    /// <param name="fontFamily">Font to use for prompting</param>
    /// <param name="fontSize">The size of the font to use in em units</param>
    /// <param name="pixelsPerDip">The pixels per density independent pixel for the current screen</param>
    /// <param name="availableWidth">The pixels available for text in the horizontal dimension</param>
    internal void LayOutLines(System.Windows.Media.FontFamily fontFamily, double fontSize, double pixelsPerDip, double availableWidth)
    {
#if DEBUG
        bool log = false;
        if (log) { LogLayoutLineAction("Start of LayOutLines()", true); }
        if (log) { LogLayoutLineAction($"fontSize={fontSize} pixelsPerDip={pixelsPerDip} availableWidth={availableWidth}"); }
#endif
        LinesOfText.Clear();

        // Bail out if we don't have content
        if (Fragments.Count == 0) return;

        int startOfLineFragIdx = 0;
        double thisFragWidth;
        double widthSoFar = 0.0;
        
        // The indent of the left margin for the current paragraph (bullet/numbering) level
        double marginIndentWidth = 0.0;
        // Used to get the next fragment out as far as the hanging indent width.
        double makeUpToTabStop = 0.0;
        // Whether this paragraph is finished
        bool paragraphComplete = false;
        // Whether the current line starts with a bullet/number requiring outdent
        bool lineStartsWithBulletNumber = false;
        // Tab stop width is the fixed size of one indent
        double tabStopWidth = fontSize * HangingIndentFactor;
        // The fragment index with which to start the subsequent line
        int nextStartOfLineFragment = -1;
        // The count of fragments in the current line
        int fragmentCount = 0;
        // The width left over from an overflow (word-wrapped) fragment
        double wordWrappedFragmentWidth = 0;
#if DEBUG
        if (log) { LogLayoutLineAction($"tabStopWidth={tabStopWidth}"); }
#endif
        //bool isBulletedOrNumbered = false;
        int consecutiveEmptyLines = 0;
        ////bool isAllHebrewFragments = true;
        bool hasAnyHebrewFragments = false;
        bool hasAnyNonHebrewFragments = false;
        bool lineComplete = false;
        for (int fragIdx = 0; fragIdx < Fragments.Count; fragIdx++)
        {
            // Pick up the next fragment and figure out what it's deal is.
            // Potential possibilities include:
            // - Just a plain old word
            // - An empty fragment with an EOL marker
            // - A word with an EOL marker
            // - A "word" that marks the start of a bullet/number paragraph
            //   NOTE: lines that are bulleted / numbered always END with an empty EOL marker fragment!
            //

            // For Debugging Line Building...
            //if (LinesOfText.Count == 21)
            //{   // Set breakpoint here...
            //    var gar = 0;
            //}

            TextFragment fragment = Fragments[fragIdx];
#if DEBUG
            if (log)
            {
                string crlf = fragment.IsLineBreak ? " CRLF" : string.Empty;
                string empty = fragment.IsEmpty ? " {}" : string.Empty;
                string ishe = fragment.HasHebrew ? " HE" : string.Empty;
                string fmt = (fragment.Bold ? "B" : string.Empty)
                           + (fragment.Italic ? "I" : string.Empty)
                           + (fragment.Underline ? "U" : string.Empty);
                string content = fragment.Content + (fragment.SuppressTrailingSpace ? "" : " ");
                LogLayoutLineAction("-------" + Environment.NewLine 
                    + $"fragIdx={fragIdx} fragContent=[{content}] (len={fragment.Content.Length}) {fmt}{crlf}{empty}{ishe}");
            }
#endif
            if (fragment.IsEmpty)
            {
                consecutiveEmptyLines++;
            }
            else
            {
                consecutiveEmptyLines = 0;
            }
            if (fragment.HasHebrew && fragment.Content.Length > 0)
            {
                hasAnyHebrewFragments = true;
            }
            else if (!fragment.HasHebrew && fragment.Content.Length > 0)
            {
                hasAnyNonHebrewFragments = true;
            }
            // Measure the width of the content of this fragment.
            thisFragWidth = MeasureFragmentWidth(fragment, fontFamily, fontSize, pixelsPerDip);
#if DEBUG
            if (log) { LogLayoutLineAction($"  -> thisFragWidth={thisFragWidth}"); }
#endif
            // Is the fragment a start of bullet/numbering indicator?
            if (fragment.NumberingId > 0 && fragment.NumberingLevel >= 0)
            {   // This is the start of a new line which has a bullet or a number.
                lineStartsWithBulletNumber = true;
                // The margin of this line will depend on the bullet/number level.
                marginIndentWidth = ((double)fragment.NumberingLevel + 1) * tabStopWidth;
                // Tab stop is the width of the extra notch less the width of the actual bullet/number content.
                makeUpToTabStop = Math.Max(tabStopWidth - thisFragWidth, 0.0);
                // Width so far is start of line + indent - outdent
                widthSoFar = marginIndentWidth - tabStopWidth;
                // If the bullet number needs padding out to the tab stop note that
                thisFragWidth += makeUpToTabStop;
#if DEBUG
                if (log)
                {
                    LogLayoutLineAction($"  >> 1st frag is bullet/number:" + Environment.NewLine
                                      + $"     marginIndentWidth={marginIndentWidth}" + Environment.NewLine
                                      + $"     makeUpToTabStop={makeUpToTabStop}" + Environment.NewLine
                                      + $"     thisFragWidth={thisFragWidth} (incl make up)");
                }
#endif
            }
            else if (fragment.IsLineBreak)
            {
                // This is the end of a line
                lineComplete = true;
                paragraphComplete = true;
                nextStartOfLineFragment = fragIdx + 1;
#if DEBUG
                if (log) { LogLayoutLineAction($"  * frag IsLineBreak => lineComplete | paragraphComplete"); }
#endif
            }
            // See how long the line is looking now...
            widthSoFar += thisFragWidth;
#if DEBUG
            if (log) { LogLayoutLineAction($"widthSoFar={widthSoFar}"); }
#endif
            // Have we overflowed the available width?
            if (widthSoFar > availableWidth)
            {
#if DEBUG
                if (log) { LogLayoutLineAction($"*** widthSoFar={widthSoFar} > availableWidth={availableWidth}"); }
#endif
                lineComplete = true;
                wordWrappedFragmentWidth = thisFragWidth;
                nextStartOfLineFragment = fragIdx;
            }
            else
            {
                fragmentCount++;
#if DEBUG
                if (log) { LogLayoutLineAction($"...included fragments={fragmentCount}"); }
#endif
            }

            if (lineComplete)
            {
                // Write the line out.
                LineOfText lineOfText = new(
                    start: startOfLineFragIdx,
                    count: fragmentCount,
                    isRightToLeft: (hasAnyHebrewFragments && !hasAnyNonHebrewFragments),
                    margin: marginIndentWidth - (lineStartsWithBulletNumber ? tabStopWidth : 0),
                    tabStop: makeUpToTabStop);
                LinesOfText.Add(lineOfText);
#if DEBUG
                if (log)
                {
                    StringBuilder stringBuilder = new();
                    for (int i = lineOfText.Start; i < (lineOfText.Start + lineOfText.Count); i++)
                    {
                        stringBuilder.Append(Fragments[i].Content + (Fragments[i].SuppressTrailingSpace ? "" : " "));
                    }
                    LogLayoutLineAction("~~~~~" + Environment.NewLine
                        + $"Line Added={LinesOfText.Count} (index={(LinesOfText.Count - 1)})" + Environment.NewLine
                        + $"+-> StartFrag= {startOfLineFragIdx} for {fragmentCount} fragments" + Environment.NewLine
                        + $"+-> margin: {marginIndentWidth - (lineStartsWithBulletNumber ? tabStopWidth : 0)}" + Environment.NewLine
                        + $"+-> tabStop: {makeUpToTabStop}" + Environment.NewLine 
                        + $"Line Content:[{stringBuilder}]" + Environment.NewLine
                        + "======="
                        );
                }
#endif
                // Reset for the next line
                startOfLineFragIdx = nextStartOfLineFragment;
                fragmentCount = (startOfLineFragIdx == fragIdx ? 1 : 0);
                lineComplete = false;
                makeUpToTabStop = 0;               // Only need the tab stop for the first line of a numbered paragraph
                lineStartsWithBulletNumber = false;
                hasAnyHebrewFragments = false;
                hasAnyNonHebrewFragments = false;
#if DEBUG
                if (log)
                {
                    LogLayoutLineAction($"Reset for next line: startOfLineFragIdx={startOfLineFragIdx}");
                }
#endif
                // Reset for the next paragraph at the end of a paragraph
                if (paragraphComplete)
                {
                    marginIndentWidth = 0;
                    makeUpToTabStop = 0;
                    paragraphComplete = false;
                }
#if DEBUG
                if (log)
                {
                    LogLayoutLineAction($"End of line: paragraphComplete={paragraphComplete} widthSoFar={widthSoFar} wordWrappedFragmentWidth={wordWrappedFragmentWidth}");
                }
#endif
                widthSoFar = marginIndentWidth + wordWrappedFragmentWidth;
                wordWrappedFragmentWidth = 0;
#if DEBUG
                if (log)
                {
                    LogLayoutLineAction($"Next line widthSoFar={widthSoFar} as starting point");
                }
#endif
            }
        }
        // Capture the last line if there's anything left...
        if (Fragments.Count - startOfLineFragIdx > 0)
        {
            LineOfText lastLineOfText = new(
                startOfLineFragIdx, 
                (Fragments.Count - startOfLineFragIdx),
                (hasAnyHebrewFragments && !hasAnyNonHebrewFragments),
                marginIndentWidth - (lineStartsWithBulletNumber ? tabStopWidth : 0),
                makeUpToTabStop);
            LinesOfText.Add(lastLineOfText);
        }
    }

    /// <summary>
    /// Measures the required width for a given TextFragment and provided display parameters
    /// </summary>
    /// <param name="fragment">The TextFragment to measure</param>
    /// <param name="fontFamily">Font to use for prompting</param>
    /// <param name="fontSize">The size of the font to use in em units</param>
    /// <param name="pixelsPerDip">The pixels per device independent pixel for the current screen</param>
    /// Omit or use string.Empty if not applicable.</param>
    /// <returns>The required width in pixels</returns>
    public static double MeasureFragmentWidth(TextFragment fragment, System.Windows.Media.FontFamily fontFamily, double fontSize, double pixelsPerDip)
    {
        string fragmentContent;
        fragmentContent = fragment.Content + (fragment.SuppressTrailingSpace ? "" : " ");
        FontStyle fontStyle = (fragment.Italic ? FontStyles.Italic : FontStyles.Normal);
        FontWeight fontWeight = (fragment.Bold ? FontWeights.Bold : FontWeights.Normal);
        var fmtTxt = new FormattedText(
            fragmentContent,
            CultureInfo.InvariantCulture,
            FlowDirection.LeftToRight,
            new Typeface(fontFamily, fontStyle, fontWeight, FontStretches.Normal),
            fontSize,
            Brushes.Black,
            pixelsPerDip);
        return fmtTxt.WidthIncludingTrailingWhitespace;
    }

    /// <summary>
    /// Checks a character to see if it falls in one of the unicode ranges for Hebrew letters and symbols
    /// </summary>
    /// <param name="input">Character to evaluate</param>
    /// <returns>True when the input character falls into one of the Hebrew unicode ranges</returns>
    private static bool IsAHebrewCharacter(char input)
    {
        // Primary Hebrew alphabet range (Aleph to Tav) plus
        // Supplemental Hebrew ranges (including diacritics and presentation forms)
        if ((input >= '\u05D0' && input <= '\u05EA') 
            || (input >= '\u0590' && input <= '\u05FF') 
            || (input >= '\uFB1D' && input <= '\uFB4F'))
        {
            return true;
        }
        return false;
    }

    /// <summary>
    /// Calculates the display for a bullet / numbering scheme given current count values
    /// </summary>
    /// <param name="scheme">The WordNumberingScheme definition that applies</param>
    /// <param name="numbers">The array of number counters (only the first 3 elements are used)</param>
    /// <returns></returns>
    private static string NumberingValue(WordNumberingScheme scheme, int[] numbers)
    {
        string result = scheme.Template;
        result = scheme.Format switch
        {
            WordNumberingFormat.Unknown => "?",
            WordNumberingFormat.Bullet => scheme.Symbol,
            WordNumberingFormat.Decimal => result.Replace("%1", numbers[0].ToString())
                                                 .Replace("%2", numbers[1].ToString())
                                                 .Replace("%3", numbers[2].ToString()),
            WordNumberingFormat.UpperLetter => result.Replace("%1", IntToLetter(numbers[0]))
                                                     .Replace("%2", IntToLetter(numbers[1]))
                                                     .Replace("%3", IntToLetter(numbers[2])),
            WordNumberingFormat.LowerLetter => result.Replace("%1", IntToLetter(numbers[0]).ToLower())
                                                     .Replace("%2", IntToLetter(numbers[1]).ToLower())
                                                     .Replace("%3", IntToLetter(numbers[2]).ToLower()),
            WordNumberingFormat.UpperRoman => result.Replace("%1", IntToRoman(numbers[0]))
                                                    .Replace("%2", IntToRoman(numbers[1]))
                                                    .Replace("%3", IntToRoman(numbers[2])),
            WordNumberingFormat.LowerRoman => result.Replace("%1", IntToRoman(numbers[0]).ToLower())
                                                    .Replace("%2", IntToRoman(numbers[1]).ToLower())
                                                    .Replace("%3", IntToRoman(numbers[2]).ToLower()),
            _ => "?",
        };
        return result;
    }

    /// <summary>
    /// Converts an integer in the range 1 to 3999 into (upper case) Roman Numerals
    /// </summary>
    /// <param name="number">The number to convert</param>
    /// <returns>Roman numerals representing the number or "?" if out of range</returns>
    private static string IntToRoman(int number)
    {
        if ((number < 0) || (number > 3999)) return "?";
        if (number < 1) return string.Empty;
        if (number >= 1000) return "M" + IntToRoman(number - 1000);
        if (number >= 900) return "CM" + IntToRoman(number - 900);
        if (number >= 500) return "D" + IntToRoman(number - 500);
        if (number >= 400) return "CD" + IntToRoman(number - 400);
        if (number >= 100) return "C" + IntToRoman(number - 100);
        if (number >= 90) return "XC" + IntToRoman(number - 90);
        if (number >= 50) return "L" + IntToRoman(number - 50);
        if (number >= 40) return "XL" + IntToRoman(number - 40);
        if (number >= 10) return "X" + IntToRoman(number - 10);
        if (number >= 9) return "IX" + IntToRoman(number - 9);
        if (number >= 5) return "V" + IntToRoman(number - 5);
        if (number >= 4) return "IV" + IntToRoman(number - 4);
        if (number >= 1) return "I" + IntToRoman(number - 1);
        return string.Empty;
    }

    /// <summary>
    /// Converts an integer in the range 1 to 26 into an (upper case) letter: (A,...,Z)
    /// </summary>
    /// <param name="number">The number to convert</param>
    /// <returns>A letter (A,...,Z) or "?" if the number is out of range</returns>
    private static string IntToLetter(int number)
    {
        if (number < 1 || number > 26)
        {
            return "?";
        }
        else
        {
            return ((char)(number + 64)).ToString();
        }
    }
}

namespace FastDocView;

/// <summary>
/// A run of blocks laid out at one page text width - a Word section. A document's sections are
/// listed in block order; section i covers blocks [FirstBlockIndex(i), FirstBlockIndex(i + 1)) and
/// the last one runs to the end. <see cref="TextWidthTwips"/> is the page width minus the left and
/// right margins and the gutter (twips, 1/1440 inch), as Word and LibreOffice lay out the text.
/// Index-time data; no file access.
/// </summary>
/// <param name="ColumnLeftTwips">Where the text column starts from the page's left edge: left margin + gutter (twips); -1 = not
/// known. Turns an object offset from the page edge into one from the column.</param>
public readonly record struct WordDocumentSection(int FirstBlockIndex, int TextWidthTwips, int ColumnLeftTwips = -1);

/// <summary>Page setup defaults and the text-width rule shared by the formats.</summary>
public static class WordDocumentPageSetup
{
    /// <summary>
    /// A document without page setup (.docx without w:pgSz / w:pgMar, .txt, web archives): Letter
    /// (8.5 in) with 1 in margins and no gutter - LibreOffice's Word import defaults
    /// (sw/source/writerfilter/dmapper/PropertyMap.cxx, SectionPropertyMap: PAPER_LETTER,
    /// m_nLeftMargin / m_nRightMargin = 1 in, m_nGutterMargin = 0).
    /// </summary>
    public const int LetterPageWidthTwips = 12240;
    public const int DefaultMarginTwips = 1440;
    public const int DefaultTextWidthTwips = LetterPageWidthTwips - 2 * DefaultMarginTwips;   // 9360 = 6.5 in

    /// <summary>The narrowest width reported; a page setup leaving less (broken margins) falls back to the default.</summary>
    public const int MinimumTextWidthTwips = 720;

    /// <summary>One section covering the whole document at the default width.</summary>
    public static IReadOnlyList<WordDocumentSection> DefaultSections { get; } = [new WordDocumentSection(0, DefaultTextWidthTwips)];

    /// <summary>
    /// Page width minus left, right margins and gutter. The gutter always narrows the text: Word adds
    /// it to the left (or inner / top-bound) margin (RTF 1.5 spec \guttersxn; LibreOffice
    /// PropertyMap.cxx subtracts GutterMargin from the left margin's text area). A gutter at the top
    /// (w:gutterAtTop) does not narrow the width; formats pass 0 then.
    /// </summary>
    public static int TextWidth(int pageWidthTwips, int leftMarginTwips, int rightMarginTwips, int gutterTwips)
    {
        long width = (long)pageWidthTwips - Math.Max(0, leftMarginTwips) - Math.Max(0, rightMarginTwips) - Math.Max(0, gutterTwips);
        return width >= MinimumTextWidthTwips && width <= 31680 * 2 ? (int)width : DefaultTextWidthTwips;
    }

    /// <summary>The text width (points) of the section holding block <paramref name="blockIndex"/>.</summary>
    public static double TextWidthPointsAt(IReadOnlyList<WordDocumentSection> sections, int blockIndex) => SectionAt(sections, blockIndex).TextWidthTwips / 20.0;

    /// <summary>Where the text column of block <paramref name="blockIndex"/>'s section starts from the page's left edge (points); -1 = not known.</summary>
    public static double ColumnLeftPointsAt(IReadOnlyList<WordDocumentSection> sections, int blockIndex)
        => SectionAt(sections, blockIndex).ColumnLeftTwips is >= 0 and var left ? left / 20.0 : -1;

    /// <summary>The left margin plus a left gutter: where the text column starts on the page (twips).</summary>
    public static int ColumnLeft(int leftMarginTwips, int gutterTwips) => Math.Max(0, leftMarginTwips) + Math.Max(0, gutterTwips);

    /// <summary>
    /// Where the gutter goes. On the left by default; on the right with a right-to-left gutter (.docx w:rtlGutter, ECMA-376
    /// 17.6.16; RTF \rtlgutter; .doc sprmSFRTLGutter 0x322A, LibreOffice ww8par6.cxx:480-487 "fRTLGutter is set if the
    /// gutter is on the right"), where it still narrows the text but no longer moves the column's start; at the top of the
    /// page with the document's gutter-at-top setting (.docx w:gutterAtTop, ECMA-376 17.15.1.49; RTF \gutterprl, LibreOffice
    /// rtfdispatchflag.cxx:1363-1364; .doc Dop iGutterPos, ww8scan.cxx:7741-7746 and ww8par6.cxx:480-483 "placed on the left
    /// unless the global dop options are to put it on top"; .odt GutterAtTop in settings.xml), where it takes nothing from the
    /// width (LibreOffice PropertyMap.cxx:760-769 moves it to the top margin).
    /// </summary>
    /// <returns>The gutter that narrows the text width, and the gutter that moves the column's start.</returns>
    public static (int WidthGutter, int LeftGutter) GutterPlacement(int gutterTwips, bool isRightToLeft, bool isAtTop)
        => isAtTop ? (0, 0) : (gutterTwips, isRightToLeft ? 0 : gutterTwips);

    private static WordDocumentSection SectionAt(IReadOnlyList<WordDocumentSection> sections, int blockIndex)
    {
        int low = 0, high = sections.Count - 1, found = -1;      // the last section starting at or before the block
        while (low <= high)
        {
            int middle = (low + high) >> 1;
            if (sections[middle].FirstBlockIndex <= blockIndex) { found = middle; low = middle + 1; } else high = middle - 1;
        }
        return found >= 0 ? sections[found] : new WordDocumentSection(0, DefaultTextWidthTwips);
    }

    /// <summary>
    /// Sections from section ends: <paramref name="lastBlockIndexes"/>[i] is the last block of section i
    /// (inclusive), <paramref name="widths"/>[i] its text width. Blocks after the last recorded end (a
    /// section whose page setup is not known yet, while the index builds) belong to the last section.
    /// </summary>
    public static IReadOnlyList<WordDocumentSection> FromSectionEnds(IReadOnlyList<int> lastBlockIndexes, IReadOnlyList<int> widths, IReadOnlyList<int>? columnLefts = null)
    {
        if (lastBlockIndexes.Count == 0) return DefaultSections;
        var sections = new List<WordDocumentSection>(lastBlockIndexes.Count);
        int first = 0;
        for (int i = 0; i < lastBlockIndexes.Count; i++)
        {
            int left = columnLefts is not null && i < columnLefts.Count ? columnLefts[i] : -1;
            if (sections.Count == 0 || sections[sections.Count - 1].TextWidthTwips != widths[i] || sections[sections.Count - 1].ColumnLeftTwips != left)
                sections.Add(new WordDocumentSection(first, widths[i], left));
            first = lastBlockIndexes[i] + 1;
        }
        return sections;
    }

    /// <summary>
    /// Sections from section starts (formats whose page setup is known where a section begins):
    /// adjacent sections with the same width are merged; the first always starts at block 0.
    /// </summary>
    public static IReadOnlyList<WordDocumentSection> FromSectionStarts(IEnumerable<WordDocumentSection> starts)
    {
        var sections = new List<WordDocumentSection>();
        foreach (WordDocumentSection section in starts)
        {
            if (sections.Count > 0 && section.FirstBlockIndex <= sections[sections.Count - 1].FirstBlockIndex) { sections[sections.Count - 1] = sections[sections.Count - 1] with { TextWidthTwips = section.TextWidthTwips, ColumnLeftTwips = section.ColumnLeftTwips }; continue; }
            if (sections.Count > 0 && sections[sections.Count - 1].TextWidthTwips == section.TextWidthTwips && sections[sections.Count - 1].ColumnLeftTwips == section.ColumnLeftTwips) continue;
            sections.Add(sections.Count == 0 ? section with { FirstBlockIndex = 0 } : section);
        }
        return sections.Count == 0 ? DefaultSections : sections;
    }

    /// <summary>
    /// A length in twips from an OOXML / ODF measure: a bare number is twips (OOXML ST_TwipsMeasure);
    /// a universal measure carries a unit (mm, cm, in, pt, pc, pi). Null when it cannot be read.
    /// </summary>
    public static int? ParseTwips(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        value = value.Trim();
        int unitStart = value.Length;
        while (unitStart > 0 && char.IsLetter(value[unitStart - 1])) unitStart--;
        if (!double.TryParse(value.Substring(0, unitStart), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double number))
            return null;
        double twips = value.Substring(unitStart).ToLowerInvariant() switch
        {
            "" => number,
            "in" => number * 1440,
            "cm" => number * 1440 / 2.54,
            "mm" => number * 1440 / 25.4,
            "pt" => number * 20,
            "pc" or "pi" => number * 240,
            _ => double.NaN,
        };
        return !double.IsNaN(twips) && !double.IsInfinity(twips) && Math.Abs(twips) < int.MaxValue ? (int)Math.Round(twips) : null;
    }
}

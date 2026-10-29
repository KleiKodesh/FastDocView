namespace FastDocView.Docx;

/// <summary>
/// The index pass's hook for list numbering: given a body-level paragraph's style id and its direct
/// numPr (numId -1 when absent), returns the next label of that list or null when the paragraph is
/// not a list item. See <see cref="DocxDocumentParts.CreateListLabelResolver"/>.
/// </summary>
public delegate DocxListLabel? DocxListLabelResolver(string? styleId, int numId, int level);

/// <summary>What separates a list label from the paragraph text (w:suff, ECMA-376 §17.9.28; tab when absent).</summary>
public enum DocxListLabelSuffix : byte { Tab, Space, Nothing }

/// <summary>
/// A list paragraph's label as counted at index time: the text ("3.", "ב."), what follows it, and the CSS of the
/// level's own run properties (lvl/rPr — font, bold, color…; empty when none). The renderer draws the label with
/// the paragraph mark's formatting plus this CSS, as Word formats numbering (§17.9.6 lvl rPr).
/// </summary>
/// <param name="LevelParagraphXml">The level's w:pPr as XML ("" when none): its indents and tab stops reach the paragraph.</param>
/// <param name="Alignment">w:lvlJc: left / start, center, right / end.</param>
public readonly record struct DocxListLabel(string Text, DocxListLabelSuffix Suffix, string LevelCss, string LevelParagraphXml = "",
                                            Html.WordListLabelAlignment Alignment = Html.WordListLabelAlignment.Start)
{
    public static Html.WordListLabelAlignment AlignmentOf(string? value) => value switch
    {
        "center" => Html.WordListLabelAlignment.Center,
        "right" or "end" => Html.WordListLabelAlignment.End,
        _ => Html.WordListLabelAlignment.Start,
    };

    public static DocxListLabelSuffix SuffixOf(string? value) => value switch
    {
        "space" => DocxListLabelSuffix.Space,
        "nothing" => DocxListLabelSuffix.Nothing,
        _ => DocxListLabelSuffix.Tab,
    };
}

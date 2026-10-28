using System.Text;
using WordDocumentRandomAccess.Html;

namespace WordDocumentRandomAccess.DocBinary;

/// <summary>
/// Renders one .doc paragraph (a CP range) to HTML: PAPX → style chain + direct paragraph sprms,
/// then CHPX runs across the paragraph's pieces for character formatting, with Word's control
/// characters mapped (tab, line break, field codes hidden, cell marks, special characters).
/// Table cells are rendered as paragraphs for now (see NOTES.md, ".doc tables").
/// </summary>
public sealed class DocBinaryParagraphHtmlRenderer
{
    private readonly DocBinaryPieceTable _pieces;
    private readonly DocBinaryFormattedDiskPages _pages;
    private readonly DocBinaryStyleSheet _styles;

    public DocBinaryParagraphHtmlRenderer(DocBinaryPieceTable pieces, DocBinaryFormattedDiskPages pages, DocBinaryStyleSheet styles)
    {
        _pieces = pieces;
        _pages = pages;
        _styles = styles;
    }

    public (WordParagraphFormatting formatting, int styleIndex) ResolveParagraph(int cpStart, int cpEnd)
    {
        uint markFc = _pieces.FileOffsetOf(Math.Max(cpStart, cpEnd - 1));
        DocBinaryFormattedDiskPages.PropertyRun? run = _pages.ParagraphRunAt(markFc);
        int styleIndex = 0;
        if (run is { GrpprlLength: >= 2 } papx) styleIndex = BitConverter.ToUInt16(papx.Page, papx.GrpprlStart);
        WordParagraphFormatting formatting = _styles.ResolveParagraphStyle(styleIndex).Clone();
        if (run is { GrpprlLength: > 2 } direct)
            DocBinaryPropertyModifiers.ApplyParagraphSprms(formatting, direct.Page, direct.GrpprlStart + 2, direct.GrpprlLength - 2);
        return (formatting, styleIndex);
    }

    public string Render(int cpStart, int cpEnd)
    {
        (WordParagraphFormatting formatting, int styleIndex) = ResolveParagraph(cpStart, cpEnd);
        int headingLevel = _styles.HeadingLevelOf(styleIndex);
        string tag = headingLevel is >= 1 and <= 6 ? "h" + headingLevel : "p";

        var html = new StringBuilder(512);
        WordHtmlText.AppendOpenTag(html, tag, formatting.ToCss(), formatting.IsInTable ? "table-cell-end" : null, formatting.IsRightToLeft);
        WordRunFormatting styleBase = _styles.ResolveParagraphStyle(styleIndex).RunFormatting;
        int fieldDepth = 0;
        foreach ((WordRunFormatting runFormatting, string text, bool isSpecial) in CharacterRuns(cpStart, cpEnd, formatting.RunFormatting, styleBase))
            AppendRunHtml(html, runFormatting, formatting.RunFormatting, text, isSpecial, ref fieldDepth);
        html.Append("</").Append(tag).Append('>');
        return html.ToString();
    }

    public string PlainText(int cpStart, int cpEnd)
    {
        var text = new StringBuilder();
        foreach (char c in _pieces.ReadText(cpStart, cpEnd)) if (c >= ' ' || c == '\t') text.Append(c == '\t' ? ' ' : c);
        return text.ToString().Trim();
    }

    private IEnumerable<(WordRunFormatting formatting, string text, bool isSpecial)> CharacterRuns(
        int cpStart, int cpEnd, WordRunFormatting paragraphBase, WordRunFormatting styleBase)
    {
        foreach ((DocBinaryPieceTable.Piece piece, int segmentStart, int segmentEnd) in _pieces.SegmentsOf(cpStart, cpEnd))
        {
            int cp = segmentStart;
            while (cp < segmentEnd)
            {
                uint fc = piece.FileOffsetOf(cp);
                DocBinaryFormattedDiskPages.PropertyRun? run = _pages.CharacterRunAt(fc);
                int runEndCp = segmentEnd;
                if (run is { } chpx && chpx.FcEnd > fc)
                    runEndCp = Math.Min(segmentEnd, cp + (int)((chpx.FcEnd - fc) / (uint)piece.BytesPerCharacter));
                if (runEndCp <= cp) runEndCp = cp + 1;

                WordRunFormatting formatting = paragraphBase.Clone();
                bool isSpecial = false;
                if (run is { GrpprlLength: > 0 } direct)
                {
                    int? characterStyle = DocBinaryPropertyModifiers.StyleIndexIn(direct.Page, direct.GrpprlStart, direct.GrpprlLength);
                    if (characterStyle is { } index && _styles[index] is { IsParagraphStyle: false } style)
                        DocBinaryPropertyModifiers.ApplyCharacterSprms(formatting, style.CharacterGrpprl, 0, style.CharacterGrpprl.Length, styleBase);
                    DocBinaryPropertyModifiers.ApplyCharacterSprms(formatting, direct.Page, direct.GrpprlStart, direct.GrpprlLength, styleBase);
                    isSpecial = DocBinaryPropertyModifiers.HasSpecialCharacterFlag(direct.Page, direct.GrpprlStart, direct.GrpprlLength);
                }
                yield return (formatting, _pieces.ReadText(cp, runEndCp), isSpecial);
                cp = runEndCp;
            }
        }
    }

    private static void AppendRunHtml(StringBuilder html, WordRunFormatting formatting, WordRunFormatting paragraphBase,
                                      string text, bool isSpecial, ref int fieldDepth)
    {
        var visible = new StringBuilder(text.Length);
        foreach (char c in text)
        {
            switch (c)
            {
                case '\u0013': fieldDepth++; continue;                          // field begin: hide the code
                case '\u0014': if (fieldDepth > 0) fieldDepth = -fieldDepth; continue;   // separator: result follows
                case '\u0015': fieldDepth = fieldDepth < 0 ? -fieldDepth - 1 : Math.Max(0, fieldDepth - 1); continue;
                case '\r': case '\u0007': case '\u000C': case '\u000E': continue; // paragraph / cell / section / column marks
                case '\t': visible.Append("<span class=\"tab\"></span>"); continue;
                case '\u000B': visible.Append("<br>"); continue;
                case '\u001E': visible.Append('‑'); continue;
                case '\u001F': visible.Append('­'); continue;
                case '\u0001': case '\u0008': if (isSpecial) visible.Append("<span class=\"unsupported\">[image]</span>"); continue;
                case '\u0002': if (isSpecial) visible.Append("<sup class=\"footnote-reference\">[*]</sup>"); continue;
                case '\u0005': continue;                                         // annotation reference
            }
            if (fieldDepth > 0) continue;                                        // inside field code
            if (c < ' ') continue;
            visible.Append(WordHtmlText.Encode(c.ToString()));
        }
        if (visible.Length == 0) return;
        string css = formatting.ToCss(paragraphBase);
        if (css.Length > 0) html.Append("<span style=\"").Append(css).Append("\">").Append(visible).Append("</span>");
        else html.Append(visible);
    }
}

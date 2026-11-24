using System.Globalization;
using System.Text;
using FastDocView.Html;

namespace FastDocView.Fields;

/// <summary>
/// Phonetic guide in .doc and RTF. Word 16 saves a w:ruby to both as an EQ field (checked: Word's SaveAs2 of a phonetic
/// guide writes "EQ \* jc0 \* "Font:Calibri" \* hps20 \o(\s\up 11(guide),base)" in .doc and in RTF's \fldinst - no
/// sprmCFRuby, no RTF ruby destination). It is recognised the way LibreOffice does (sw/source/filter/ww8/ww8par5.cxx:2746
/// Read_SubF_Ruby: code starts with \*, \* jcN / hpsN / "Font:name", then \o with \s\up N(guide),base) and drawn as HTML
/// ruby, the same markup as .docx's w:ruby. Anything else stays a general EQ (MathML).
/// Word writes the guide and base unescaped, so brackets, commas and backslashes in them land in the code as they are. Word's
/// own EQ parser (test-data/word-ruby/ruby-escapes-v1 and ruby-escapes-probe-v1, Word 16 PDFs) then ends the guide at the
/// first comma (with none, at its first ")" - ";" separates nothing), takes everything after it as the base, drops every ")"
/// and keeps "(" and "\" as text: a backslash escapes nothing here ("a\,b" still splits, "k\\l" draws both backslashes).
/// LibreOffice instead cuts the guide at the first ")" and the base at the last one. A font name Word writes quoted; an
/// unquoted one runs to the next switch, so Font:Times New Roman is one name (LibreOffice keeps only the first word).
/// </summary>
public static class WordFieldRuby
{
    /// <summary>
    /// Stands for a note reference inside a field code (.doc 0x02, an RTF {\footnote} in \fldinst): the readers put it in the
    /// code, the renderers swap it for the note's mark in the field's HTML (a private-use character, never document text here).
    /// </summary>
    public const char NoteMarkPlaceholder = '';
    /// <summary><see cref="NoteMarkPlaceholder"/> as a string (what a reader routes for a note reference inside fields).</summary>
    public static readonly string NoteMarkPlaceholderText = NoteMarkPlaceholder.ToString();

    /// <summary>
    /// The class of the span that holds a guide's rest after its comma on the base line (all three formats): the app's search
    /// highlight reads it right after the guide, with the comma between, as Word's Find reads such a guide (see ToHtml).
    /// </summary>
    public const string GuideRestClass = "ruby-guide-rest";

    /// <summary>The ruby HTML, or null when the code is not Word's phonetic-guide form.</summary>
    public static string? ToHtml(string equationCode) => ToHtml(equationCode, out _);

    /// <summary>
    /// The ruby HTML (null when the code is not Word's phonetic-guide form) and its search text: what Word 16's Find matches
    /// in it - the guide, then the base, as one run between the text around the field ("A {EQ ...(gd),base} Z" finds "A gd",
    /// "gdbase" and "base Z", not "basegd"; hidden-Word probe 2026-11-24). Word's Range.Text and its plain-text save give
    /// nothing for any EQ; the search follows Find. A guide with a comma before its ")" (drawn split: head over the rest, a comma
    /// and the base) is read as written, head + "," + rest, and its base not at all (hidden Word 16 Find on ruby-edges-v1 and
    /// ruby-docx-commas-v1, 2026-11-24: "g,3", "g,3 after", "g,1,2" found; "base", ",base", "3,base" not).
    /// </summary>
    public static string? ToHtml(string equationCode, out string? searchText)
    {
        searchText = null;
        string code = equationCode.Trim();
        if (!code.StartsWith(@"\*", StringComparison.Ordinal)) return null;
        int justification = 0, halfPoints = 0;
        string fontName = "";
        int position = 0;
        // the \* options before \o
        while (position < code.Length)
        {
            while (position < code.Length && char.IsWhiteSpace(code[position])) position++;
            if (!code.AsSpan(position).StartsWith(@"\*".AsSpan(), StringComparison.Ordinal)) break;
            position += 2;
            while (position < code.Length && char.IsWhiteSpace(code[position])) position++;
            string option;
            if (position < code.Length && code[position] == '"')
            {
                int quoteEnd = code.IndexOf('"', position + 1);
                if (quoteEnd < 0) return null;
                option = code.Substring(position + 1, quoteEnd - position - 1);
                position = quoteEnd + 1;
            }
            else
            {
                int start = position;
                while (position < code.Length && !char.IsWhiteSpace(code[position]) && code[position] != '\\') position++;
                option = code.Substring(start, position - start);
                if (option.StartsWith("Font:", StringComparison.OrdinalIgnoreCase))
                {
                    // an unquoted font name with spaces runs to the next switch
                    while (position < code.Length && code[position] != '\\') position++;
                    option = code.Substring(start, position - start).TrimEnd();
                }
            }
            if (option.StartsWith("jc", StringComparison.OrdinalIgnoreCase)) int.TryParse(option.Substring(2), NumberStyles.Integer, CultureInfo.InvariantCulture, out justification);
            else if (option.StartsWith("hps", StringComparison.OrdinalIgnoreCase)) int.TryParse(option.Substring(3), NumberStyles.Integer, CultureInfo.InvariantCulture, out halfPoints);
            else if (option.StartsWith("Font:", StringComparison.OrdinalIgnoreCase)) fontName = option.Substring(5).Trim();
        }
        string rest = code.Substring(position);
        if (!rest.StartsWith(@"\o", StringComparison.OrdinalIgnoreCase)) return null;
        // \s\up must be \o's first element ("\o(" or "\o\aX(", then \s\up): any other \o is an equation
        int overstrikeOpen = rest.IndexOf('(');
        if (overstrikeOpen < 0 || rest.IndexOf(' ', 0, overstrikeOpen) >= 0) return null;
        string first = rest.Substring(overstrikeOpen + 1).TrimStart();
        if (!first.StartsWith(@"\s\up", StringComparison.OrdinalIgnoreCase)) return null;
        string part = first.Substring(5);
        int open = part.IndexOf('(');
        if (open < 0) return null;
        int comma = part.IndexOf(',', open + 1);
        int last = part.LastIndexOf(')');      // \o's closing bracket: nothing after it is base text
        if (comma < 0)
        {
            // no comma: ";" is not a separator for Word (ruby-edges-semicolon-v1.pdf: "(g;6);b;6)" prints "g;6" over ";b;6",
            // LibreOffice would split at the ";"): the guide ends at its ")", the rest up to the last ")" is the base
            int close = part.IndexOf(')', open + 1);
            if (close < 0 || last <= close) return null;
            string onlyGuide = part.Substring(open + 1, close - open - 1);
            string onlyBase = part.Substring(close + 1, last - close - 1).Replace(")", "");
            if (onlyGuide.Length == 0 || onlyBase.Length == 0) return null;
            comma = close;   // the rest below reads guide [open+1, comma) and base (comma, last) alike
        }
        if (last <= comma) return null;
        // Word drops every ")" from the guide and the base (they close its groups) and keeps everything else as text
        string guide = part.Substring(open + 1, comma - open - 1).Replace(")", "");
        string baseText = part.Substring(comma + 1, last - comma - 1).Replace(")", "");
        if (guide.Length == 0 || baseText.Length == 0) return null;
        // a comma inside the guide (before its first ")"): the guide's rest, up to that ")", is drawn on the base line in the
        // guide's font and size (ruby-docx-commas-v1.doc.pdf: guide "g,1,2" over base "base" prints "1,2" in Courier New, then
        // ",base" in the base font)
        int firstClose = part.IndexOf(')', open + 1);
        string guideRest = "";
        if (firstClose > comma)
        {
            guideRest = part.Substring(comma + 1, firstClose - comma - 1);
            baseText = part.Substring(firstClose, last - firstClose).Replace(")", "");
            if (guideRest.Length + baseText.Length == 0) return null;
        }

        searchText = (guideRest.Length == 0 ? guide + baseText : guide + "," + guideRest).Replace(NoteMarkPlaceholderText, "");
        var font = new StringBuilder();
        if (halfPoints > 0) font.Append("font-size:").Append((halfPoints / 2m).ToString(CultureInfo.InvariantCulture)).Append("pt;");
        if (fontName.Length > 0) font.Append("font-family:'").Append(WordHtmlText.CssFontName(fontName)).Append("';");
        string style = font + AlignCss(justification);
        string baseHtml = (guideRest.Length == 0 ? "" : "<span class=\"" + GuideRestClass + "\"" + (font.Length > 0 ? " style=\"" + font + "\"" : "") + ">" + WordHtmlText.Encode(guideRest) + "</span>")
                          + WordHtmlText.Encode(baseText);
        return "<ruby>" + baseHtml + "<rt" + (style.Length > 0 ? " style=\"" + style + "\"" : "") + ">"
               + WordHtmlText.Encode(guide) + "</rt></ruby>";
    }

    /// <summary>
    /// The guide's alignment as CSS, by EQ's jcN (LibreOffice ww8par5.cxx:2832-2851): 0 centre (the browser default, ""),
    /// 1 distribute letters (0:1:0), 2 distribute space (1:2:1), 3 left, 4 right.
    /// </summary>
    public static string AlignCss(int justification) => justification switch
    {
        1 => "ruby-align:space-between;", 2 => "ruby-align:space-around;", 3 => "ruby-align:start;", 4 => "ruby-align:end;", _ => "",
    };

    /// <summary>.docx w:rubyAlign (ST_RubyAlign, the same five as jc; LibreOffice DomainMapper.cxx:3905 convertRubyAlign) as CSS.</summary>
    public static string AlignCss(string? rubyAlign) => AlignCss(rubyAlign switch
    {
        "distributeLetter" => 1, "distributeSpace" => 2, "left" => 3, "right" => 4, _ => 0,
    });
}

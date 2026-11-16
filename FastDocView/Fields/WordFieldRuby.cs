using System.Globalization;
using System.Text;
using FastDocView.Html;

namespace FastDocView.Fields;

/// <summary>
/// Phonetic guide in .doc and RTF. Word 16 saves a w:ruby to both as an EQ field (checked: Word's SaveAs2 of a phonetic
/// guide writes "EQ \* jc0 \* "Font:Calibri" \* hps20 \o(\s\up 11(guide),base)" in .doc and in RTF's \fldinst - no
/// sprmCFRuby, no RTF ruby destination). It is recognised the way LibreOffice does (sw/source/filter/ww8/ww8par5.cxx:2746
/// Read_SubF_Ruby: code starts with \*, \* jcN / hpsN / "Font:name", then \o with \s\up N(guide) and the base after the
/// comma up to the last ")" - Word disallows brackets inside) and drawn as HTML ruby, the same markup as .docx's w:ruby.
/// Anything else stays a general EQ (MathML).
/// </summary>
public static class WordFieldRuby
{
    /// <summary>The ruby HTML, or null when the code is not Word's phonetic-guide form.</summary>
    public static string? ToHtml(string equationCode)
    {
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
            }
            if (option.StartsWith("jc", StringComparison.OrdinalIgnoreCase)) int.TryParse(option.Substring(2), NumberStyles.Integer, CultureInfo.InvariantCulture, out justification);
            else if (option.StartsWith("hps", StringComparison.OrdinalIgnoreCase)) int.TryParse(option.Substring(3), NumberStyles.Integer, CultureInfo.InvariantCulture, out halfPoints);
            else if (option.StartsWith("Font:", StringComparison.OrdinalIgnoreCase)) fontName = option.Substring(5).Trim();
        }
        string rest = code.Substring(position);
        if (!rest.StartsWith(@"\o", StringComparison.OrdinalIgnoreCase)) return null;
        int up = rest.IndexOf(@"\s\up", StringComparison.OrdinalIgnoreCase);
        if (up < 0) return null;
        string part = rest.Substring(up + 5);
        int open = part.IndexOf('('), close = part.IndexOf(')');
        if (open < 0 || close < open) return null;
        string guide = part.Substring(open + 1, close - open - 1);
        int comma = part.IndexOf(',', close);
        if (comma < 0) comma = part.IndexOf(';', close);
        int last = part.LastIndexOf(')');
        if (comma < 0 || last <= comma) return null;
        string baseText = part.Substring(comma + 1, last - comma - 1);
        if (guide.Length == 0 || baseText.Length == 0) return null;

        var style = new StringBuilder();
        if (halfPoints > 0) style.Append("font-size:").Append((halfPoints / 2m).ToString(CultureInfo.InvariantCulture)).Append("pt;");
        if (fontName.Length > 0) style.Append("font-family:'").Append(WordHtmlText.CssFontName(fontName)).Append("';");
        style.Append(AlignCss(justification));
        return "<ruby>" + WordHtmlText.Encode(baseText) + "<rt" + (style.Length > 0 ? " style=\"" + style + "\"" : "") + ">"
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

using System.Xml.Linq;

namespace FastDocView.DocxHtml;

/// <summary>
/// A .docx w:ruby's guide (w:rt) and base (w:rubyBase) as Word prints them. Word draws a phonetic guide through its EQ
/// rules even from .docx, so brackets and commas in the text change it (Word 16 PDFs of test-data/word-ruby/
/// ruby-escapes-v1.docx and ruby-docx-commas-v1.docx). LibreOffice passes the text through as written
/// (sw/source/writerfilter/dmapper/DomainMapper.cxx:4587 SetRubyText, :3903 PROP_RUBY_TEXT), so Word is the reference:
/// - guide and base each end at their first ")" (the rest is not drawn); "(" is kept;
/// - a comma in the guide ends it there, and the base line then shows the guide's rest, a comma in the base's font, and the
///   guide's rest again - the base's own text is not drawn (base "base", guide "g,1,2": "g" over "1,2" "," "1,2", the rest
///   in the guide's font).
/// Only the elements' w:t text changes; a ruby without ")" or "," is returned as it is (no copy).
/// </summary>
public static class DocxRubyText
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public static (XElement Base, XElement Guide) AsWordPrints(XElement rubyBase, XElement guide)
    {
        string baseText = TextOf(rubyBase), guideText = TextOf(guide);
        if (baseText.IndexOf(')') < 0 && guideText.IndexOfAny([')', ',']) < 0) return (rubyBase, guide);
        XElement newBase = Keep(rubyBase, 0, CutAt(baseText));
        int guideEnd = CutAt(guideText);
        int comma = guideText.IndexOf(',', 0, guideEnd);
        if (comma < 0) return (newBase, Keep(guide, 0, guideEnd));
        XElement head = Keep(guide, 0, comma);
        XElement rest = Keep(guide, comma + 1, guideEnd);
        // the comma between the two copies is drawn in the base's run formatting
        XElement? baseRun = Shown(rubyBase).FirstOrDefault()?.Parent;
        var commaRun = new XElement(W + "r", baseRun?.Element(W + "rPr") is { } properties ? new XElement(properties) : null, new XElement(W + "t", ","));
        var line = new XElement(rubyBase.Name, rest.Nodes().Select(Copy), commaRun, rest.Nodes().Select(Copy));
        return (line, head);
    }

    private static int CutAt(string text) => text.IndexOf(')') is int bracket and >= 0 ? bracket : text.Length;

    private static string TextOf(XElement container) => string.Concat(Shown(container).Select(t => t.Value));

    /// <summary>The drawn w:t elements: not inside a tracked move's source (w:moveFrom, hidden in Word's final view).</summary>
    private static IEnumerable<XElement> Shown(XElement container)
        => container.Descendants(W + "t").Where(t => !t.Ancestors().TakeWhile(a => a != container).Any(a => a.Name == W + "moveFrom"));

    /// <summary>A copy of the container whose w:t text keeps only characters [start, end) of its whole text.</summary>
    private static XElement Keep(XElement container, int start, int end)
    {
        var copy = new XElement(container);
        int offset = 0;
        foreach (XElement text in Shown(copy).ToList())
        {
            string value = text.Value;
            int from = Math.Clamp(start - offset, 0, value.Length), to = Math.Clamp(end - offset, 0, value.Length);
            offset += value.Length;
            text.Value = to > from ? value.Substring(from, to - from) : "";
        }
        return copy;
    }

    private static XNode Copy(XNode node) => node is XElement element ? new XElement(element) : node is XText text ? new XText(text) : node;
}

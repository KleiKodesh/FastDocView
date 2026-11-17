using System.Xml.Linq;

namespace FastDocView.DocxHtml;

/// <summary>
/// A .docx w:ruby's guide (w:rt) and base (w:rubyBase) as Word prints them. Word draws a phonetic guide through its EQ rules
/// even from .docx (Word 16 PDFs of test-data/word-ruby/ruby-escapes-v1.docx, ruby-docx-commas-v1.docx and the hand-written
/// ruby-edges-v1.docx), the same rules as a .doc's EQ field (Fields.WordFieldRuby):
/// - every ")" in the guide and the base is not drawn; the text after it is ("ab)c", a tab, "d" prints "abc d"); tabs, note
///   marks and bookmarks stay where they are;
/// - a comma in the guide ends it there, and the guide's rest after the comma is drawn at the start of the base line in the
///   guide's runs, then a comma in the base's run formatting, then the base itself (guide "g,3" over base "base" + a note
///   mark prints "g" over "3,base" + the mark). Word's UI writes such a ruby with the base already replaced by the rest
///   (guide "g,1" over base "1" prints "1,1").
/// LibreOffice passes the text through as written (sw/source/writerfilter/dmapper/DomainMapper.cxx:4587 SetRubyText, :3903
/// PROP_RUBY_TEXT), so Word is the reference. Search and GetText read the base the same way (DocxSearchTextScanner and
/// WordDocumentSearch feed <see cref="SearchRest"/> and a comma first, then the base without its ")").
/// A ruby without ")" or "," is returned as it is (no copy). Every copied w:t is mapped to its original so paragraph-end
/// handling (trailing spaces) still sees where it sits.
/// </summary>
public static class DocxRubyText
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>The guide's rest after its first comma, without ")" - drawn and searched at the start of the base - or null.</summary>
    public static string? SearchRest(string guideText)
    {
        int comma = guideText.IndexOf(',');
        return comma < 0 ? null : guideText.Substring(comma + 1).Replace(")", "");
    }

    /// <summary>The base line and the guide to draw; <paramref name="originals"/> gets every copied w:t's original.</summary>
    public static (XElement Base, XElement Guide) AsWordPrints(XElement rubyBase, XElement guide, Dictionary<XElement, XElement> originals)
    {
        string baseText = TextOf(rubyBase), guideText = TextOf(guide);
        if (baseText.IndexOf(')') < 0 && guideText.IndexOfAny([')', ',']) < 0) return (rubyBase, guide);
        XElement newBase = Keep(rubyBase, 0, int.MaxValue, originals);
        int comma = guideText.IndexOf(',');
        if (comma < 0) return (newBase, Keep(guide, 0, int.MaxValue, originals));
        XElement head = Keep(guide, 0, comma, originals);
        XElement rest = Keep(guide, comma + 1, int.MaxValue, originals);
        // the comma after the rest is drawn in the base's run formatting (its first drawn run)
        XElement? baseRun = Shown(rubyBase).FirstOrDefault()?.Parent;
        var commaRun = new XElement(W + "r", baseRun?.Element(W + "rPr") is { } properties ? new XElement(properties) : null, new XElement(W + "t", ","));
        // detached first, so the line takes these very nodes (a parented node would be copied and lose its mapping)
        List<XNode> restNodes = rest.Nodes().ToList(), baseNodes = newBase.Nodes().ToList();
        rest.RemoveNodes();
        newBase.RemoveNodes();
        return (new XElement(rubyBase.Name, restNodes, commaRun, baseNodes), head);
    }

    private static string TextOf(XElement container) => string.Concat(Shown(container).Select(t => t.Value));

    /// <summary>The drawn w:t elements: not inside a tracked move's source (w:moveFrom, hidden in Word's final view).</summary>
    private static IEnumerable<XElement> Shown(XElement container)
        => container.Descendants(W + "t").Where(t => !t.Ancestors().TakeWhile(a => a != container).Any(a => a.Name == W + "moveFrom"));

    /// <summary>
    /// A copy of the container whose drawn text keeps only characters [start, end) of its whole drawn text, without ")";
    /// everything else (tabs, note marks, bookmarks) is kept.
    /// </summary>
    private static XElement Keep(XElement container, int start, int end, Dictionary<XElement, XElement> originals)
    {
        var copy = new XElement(container);
        List<XElement> originalTexts = Shown(container).ToList(), copiedTexts = Shown(copy).ToList();
        int offset = 0;
        for (int i = 0; i < copiedTexts.Count; i++)
        {
            XElement text = copiedTexts[i];
            string value = text.Value;
            int from = Math.Min(Math.Max(start - offset, 0), value.Length), to = Math.Min(Math.Max(end - offset, 0), value.Length);
            offset += value.Length;
            text.Value = (to > from ? value.Substring(from, to - from) : "").Replace(")", "");
            originals[text] = originalTexts[i];
        }
        return copy;
    }
}

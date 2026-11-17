using System.Text;

namespace FastDocView.Html;

/// <summary>
/// Floats placed from their paragraph's top (vertical position relative to the paragraph: DrawingML positionV paragraph, VML
/// text, .doc by / posrelv 2, RTF \shpbypara / posrelv 2) are lifted to the start of their paragraph: Word measures the offset
/// from the paragraph's top line, wherever the anchor sits in it, while a CSS float starts on the line of its place in the text
/// (Word 16's PDF of test-data/word-shapes/floating-v1: tblate, anchored late in its paragraph). Line-relative floats stay on
/// their line. <see cref="WordPictureHtml.AppendPlacement"/> marks the frames; this runs once per rendered block.
/// </summary>
internal static class WordParagraphFloats
{
    /// <summary>The style declaration marking a float to lift (a custom property: no effect on layout).</summary>
    public const string Marker = "--word-float:paragraph;";
    private const string Done = "--word-float:lifted;   ";

    /// <summary>The block's HTML with every marked float moved to just after its paragraph's opening tag (the &lt;p&gt; or heading that holds it, not across a text box or shape).</summary>
    public static string Hoist(string html)
    {
        if (html.IndexOf(Marker, StringComparison.Ordinal) < 0) return html;
        var text = new StringBuilder(html);
        for (int marker = IndexOf(text, Marker, 0); marker >= 0; marker = IndexOf(text, Marker, marker + 1))
        {
            int start = LastIndexOf(text, "<", marker);
            int end = FrameEnd(text, start);
            int paragraph = ParagraphOpenEnd(text, start);
            ReplaceAt(text, marker, Done);   // same length: positions stay
            if (end < 0 || paragraph < 0 || paragraph == start) continue;
            string frame = text.ToString(start, end - start);
            text.Remove(start, end - start);
            text.Insert(paragraph, frame);
            marker = paragraph;
        }
        return text.Replace(Done, "").ToString();   // the markers go: the page sees only the frame's own style
    }

    private static void ReplaceAt(StringBuilder text, int at, string value)
    {
        for (int i = 0; i < value.Length; i++) text[at + i] = value[i];
    }

    /// <summary>The end (exclusive) of the element opening at <paramref name="start"/>: an &lt;img&gt;'s tag, a &lt;span&gt; through its matching close.</summary>
    private static int FrameEnd(StringBuilder text, int start)
    {
        if (start < 0) return -1;
        if (StartsWith(text, start, "<img")) return IndexOf(text, ">", start) is var close and >= 0 ? close + 1 : -1;
        if (!StartsWith(text, start, "<span")) return -1;
        int depth = 0;
        for (int at = start; at < text.Length; at++)
        {
            if (text[at] != '<') continue;
            if (StartsWith(text, at, "<span")) depth++;
            else if (StartsWith(text, at, "</span>") && --depth == 0) return at + 7;
        }
        return -1;
    }

    /// <summary>
    /// Where the content of the paragraph holding <paramref name="frame"/> starts (just after its &lt;p ...&gt;), or -1 when
    /// the frame is not directly in a paragraph's text (inside a text box or shape, whose paragraphs are spans).
    /// </summary>
    private static int ParagraphOpenEnd(StringBuilder text, int frame)
    {
        // a paragraph is <p> or a heading <h1>..<h6> (the renderers' paragraph tags)
        int open = -1, nameLength = 0;
        for (int at = frame - 1; at >= 0 && open < 0; at--)
        {
            if (text[at] != '<' || at + 3 >= text.Length) continue;
            int length = text[at + 1] == 'p' ? 1 : text[at + 1] == 'h' && text[at + 2] is >= '1' and <= '6' ? 2 : 0;
            if (length > 0 && text[at + 1 + length] is ' ' or '>') { open = at; nameLength = length; }
        }
        if (open < 0 || IndexOf(text, "</" + text.ToString(open + 1, nameLength) + ">", open) is var close and >= 0 && close < frame) return -1;
        int contentStart = IndexOf(text, ">", open) + 1;
        if (contentStart <= 0 || contentStart > frame) return -1;
        // the spans still open between the paragraph and the frame: runs may be left, a text box or shape may not
        var openSpans = new Stack<bool>();
        for (int at = contentStart; at < frame; at++)
        {
            if (text[at] != '<') continue;
            if (StartsWith(text, at, "<span"))
            {
                int tagEnd = IndexOf(text, ">", at);
                string tag = tagEnd < 0 ? "" : text.ToString(at, tagEnd - at);
                openSpans.Push(tag.Contains("text-box") || tag.Contains("word-shape"));
            }
            else if (StartsWith(text, at, "</span>") && openSpans.Count > 0) openSpans.Pop();
        }
        return openSpans.Contains(true) ? -1 : contentStart;
    }

    private static bool StartsWith(StringBuilder text, int at, string value)
    {
        if (at < 0 || at + value.Length > text.Length) return false;
        for (int i = 0; i < value.Length; i++) if (text[at + i] != value[i]) return false;
        return true;
    }

    private static int IndexOf(StringBuilder text, string value, int from)
    {
        for (int at = Math.Max(0, from); at + value.Length <= text.Length; at++) if (StartsWith(text, at, value)) return at;
        return -1;
    }

    private static int LastIndexOf(StringBuilder text, string value, int before)
    {
        for (int at = Math.Min(before, text.Length - value.Length); at >= 0; at--) if (StartsWith(text, at, value)) return at;
        return -1;
    }
}

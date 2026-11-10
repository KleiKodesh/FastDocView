using System.Text;
using System.Xml.Linq;
using FastDocView.Html;

namespace FastDocView.Rtf;

/// <summary>
/// An RTF math group ({\mmath ...}, RTF 1.9.1 "Office Math" control words) rebuilt as OMML, for
/// <see cref="WordOfficeMathHtml"/>. RTF spells OMML one to one: a group {\mX ...} is the element m:X (\moMath, \mf, \mnum,
/// \msSubSup...), a property group holding only text ({\mchr ∑}, {\mbegChr [}, {\mtype lin}, {\mlimLoc undOvr}...) is that
/// element with m:val = the text, and a few properties are control words with a number: \mscrN (m:scr), \mstyN (m:sty),
/// \mjcN (m:jc), \mnor and \mlit flags. This is how LibreOffice's RTF import reads it (rtfdispatchdestination.cxx:537-578:
/// every math keyword opens its OMML element; rtfdocumentimpl.cxx:3316-3370: the value destinations become m:val), and the
/// numbers are Word's (math-v1.rtf, saved by Word 16: \mscr0..5 = roman, script, fraktur, double-struck, sans-serif,
/// monospace; \msty0..3 = p, b, i, bi; \mjc3 = left, \mjc4 = right, \mdefJc1 = centerGroup). LibreOffice reads neither
/// \mscr / \msty nor \mjc (rtfdispatchvalue.cxx has no case for them). {\mmathPict} (the picture for other readers) is
/// skipped; text outside an {\mr} is a run of its own; \cfN gives a run its colour (w:color).
/// </summary>
internal static class RtfOfficeMath
{
    private static readonly XNamespace M = WordOfficeMathHtml.MathNamespace;
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>The property elements whose content is their m:val (ECMA-376 22.1.2: ST_Char, ST_OnOff, ST_TopBot... values).</summary>
    private static readonly HashSet<string> ValueProperties = new(StringComparer.Ordinal)
    {
        "chr", "begChr", "sepChr", "endChr", "pos", "vertJc", "type", "limLoc", "grow", "show", "degHide", "subHide", "supHide",
        "hideTop", "hideBot", "hideLeft", "hideRight", "strikeH", "strikeV", "strikeBLTR", "strikeTLBR", "count", "mcJc", "zeroWid",
        "zeroAsc", "zeroDesc", "transp", "opEmu", "noBreak", "diff", "baseJc", "maxDist", "objDist", "plcHide", "shp", "argSz",
        "rSp", "rSpRule", "cGp", "cGpRule", "cSp", "alnScr", "aln", "brk", "jc",
    };

    private static readonly string[] Scripts = ["roman", "script", "fraktur", "double-struck", "sans-serif", "monospace"];
    private static readonly string[] Styles = ["p", "b", "i", "bi"];
    private static readonly string?[] Justifications = [null, "centerGroup", "center", "left", "right"];

    private sealed class Frame
    {
        public XElement? Element;          // the math element this group opened (null: a plain formatting group)
        public bool Skips;                 // {\mmathPict}, {\*\unknown}: nothing inside is read
        public int Color = -1;             // \cfN in force
        public int UnicodeSkip = 1;        // \ucN in force
        public StringBuilder? Value;       // a value property's text
    }

    /// <summary>The m:oMathPara / m:oMath of a captured math group (its bytes from after the opening '{'), or null.</summary>
    public static XElement? ToOmml(ReadOnlySpan<byte> group, int codePage, Func<int, string?> colorOf)
    {
        var root = new XElement(M + "root");
        var frames = new List<Frame> { new() { Element = root } };
        XElement? currentRun = null;           // the m:r taking text now
        Frame? runFrame = null;                // the group that opened it ({\mr}); null: a bare-text run
        int pendingSkip = 0;                   // fallback characters still to skip after \uN
        bool starred = false;                  // \* just seen: an unknown destination is skipped
        bool groupStart = true;                // at the start of a group (its first control word may name it)

        Frame Top() => frames[frames.Count - 1];
        XElement Container()
        {
            for (int i = frames.Count - 1; i >= 0; i--) if (frames[i].Element is { } e) return e;
            return root;
        }
        void AddText(string text)
        {
            Frame top = Top();
            if (top.Skips || text.Length == 0) return;
            for (int i = frames.Count - 1; i >= 0; i--)
                if (frames[i].Value is { } value) { value.Append(text); return; }
                else if (frames[i].Element is not null) break;
            if (currentRun is null)
            {
                XElement container = Container();
                if (container == root || container.Name.LocalName.EndsWith("Pr", StringComparison.Ordinal)) return;   // text outside an equation / in properties
                currentRun = new XElement(M + "r");
                ApplyColor(currentRun);
                container.Add(currentRun);
                runFrame = null;
            }
            if (currentRun.Elements(M + "t").LastOrDefault() is { } t && t == currentRun.LastNode) t.Value += text;
            else currentRun.Add(new XElement(M + "t", text));
        }
        void ApplyColor(XElement run)
        {
            for (int i = frames.Count - 1; i >= 0; i--)
                if (frames[i].Color >= 0)
                {
                    if (colorOf(frames[i].Color) is { } hex) run.Add(new XElement(W + "rPr", new XElement(W + "color", new XAttribute(W + "val", hex))));
                    return;
                }
        }
        XElement RunProperties(XElement run)
        {
            if (run.Element(M + "rPr") is { } existing) return existing;
            var properties = new XElement(M + "rPr");
            run.AddFirst(properties);
            return properties;
        }

        int position = 0;
        while (position < group.Length)
        {
            byte b = group[position++];
            if (b == '{')
            {
                frames.Add(new Frame { Skips = Top().Skips, Color = -1, UnicodeSkip = Top().UnicodeSkip });
                groupStart = true;
                starred = false;
                continue;
            }
            if (b == '}')
            {
                Frame closed = Top();
                if (frames.Count > 1) frames.RemoveAt(frames.Count - 1);
                if (closed.Value is { } value && closed.Element is { } property) property.SetAttributeValue(M + "val", value.ToString());
                if (closed == runFrame || closed.Element is not null) { currentRun = null; runFrame = null; }
                groupStart = false;
                continue;
            }
            if (b is (byte)'\r' or (byte)'\n') continue;
            if (b != '\\')
            {
                if (pendingSkip > 0) { pendingSkip--; continue; }
                int start = position - 1;
                while (position < group.Length && group[position] is not ((byte)'\\' or (byte)'{' or (byte)'}' or (byte)'\r' or (byte)'\n')) position++;
                AddText(Decode(group.Slice(start, position - start), codePage));
                groupStart = false;
                continue;
            }
            if (position >= group.Length) break;
            byte next = group[position];
            if (next == '*') { position++; starred = true; continue; }
            if (next == '\'')
            {
                if (position + 2 >= group.Length) break;
                int hex = HexValue(group[position + 1]) * 16 + HexValue(group[position + 2]);
                position += 3;
                if (pendingSkip > 0) { pendingSkip--; continue; }
                if (hex >= 0) AddText(Decode([(byte)hex], codePage));
                continue;
            }
            if (!IsLetter(next))
            {
                position++;
                if (pendingSkip > 0) { pendingSkip--; continue; }
                if (next is (byte)'\\' or (byte)'{' or (byte)'}') AddText(((char)next).ToString());
                else if (next == '~') AddText(" ");
                else if (next == '-') { }
                else if (next == '_') AddText("‑");
                continue;
            }
            int nameStart = position;
            while (position < group.Length && IsLetter(group[position])) position++;
            string word = Encoding.ASCII.GetString(group.Slice(nameStart, position - nameStart).ToArray());
            bool hasParameter = false, negative = false;
            long parameter = 0;
            if (position < group.Length && group[position] == '-') { negative = true; position++; }
            while (position < group.Length && group[position] is >= (byte)'0' and <= (byte)'9') { hasParameter = true; parameter = Math.Min(parameter * 10 + (group[position] - '0'), int.MaxValue); position++; }
            if (negative) parameter = -parameter;
            if (position < group.Length && group[position] == ' ') position++;
            bool wasGroupStart = groupStart;
            groupStart = false;
            Frame frame = Top();
            if (word == "bin" && hasParameter) { position = (int)Math.Min(group.Length, position + Math.Max(0, parameter)); continue; }
            if (frame.Skips) continue;
            pendingSkip = 0;
            switch (word)
            {
                case "u":
                {
                    int code = (int)(parameter < 0 ? parameter + 65536 : parameter);     // \uN: a signed 16-bit value; surrogate halves pair up in the text
                    AddText(((char)Math.Max(0, Math.Min(code, 0xFFFF))).ToString());
                }
                    pendingSkip = frame.UnicodeSkip;
                    continue;
                case "uc": frame.UnicodeSkip = (int)Math.Max(0, parameter); continue;
                case "cf": frame.Color = hasParameter ? (int)parameter : 0; continue;
                case "tab": AddText(" "); continue;
                case "mmath": continue;                                     // the container: its equation follows
                case "mmathPict": case "mmathPr": frame.Skips = true; continue;
                case "mscr" when currentRun is not null && hasParameter && parameter is >= 0 and < 6:
                    RunProperties(currentRun).Add(new XElement(M + "scr", new XAttribute(M + "val", Scripts[parameter]))); continue;
                case "msty" when currentRun is not null && hasParameter && parameter is >= 0 and < 4:
                    RunProperties(currentRun).Add(new XElement(M + "sty", new XAttribute(M + "val", Styles[parameter]))); continue;
                case "mnor" when currentRun is not null: RunProperties(currentRun).Add(new XElement(M + "nor")); continue;
                case "mlit" when currentRun is not null: RunProperties(currentRun).Add(new XElement(M + "lit")); continue;
                case "mjc" when hasParameter && parameter is >= 1 and <= 4 && Container().Name.LocalName == "oMathParaPr":
                    Container().Add(new XElement(M + "jc", new XAttribute(M + "val", Justifications[parameter]!))); continue;
            }
            if (!wasGroupStart || word.Length < 2 || word[0] != 'm')
            {
                if (wasGroupStart && starred) frame.Skips = true;       // {\*\unknown ...}: an ignorable destination
                starred = false;
                continue;
            }
            starred = false;
            // {\mX ...}: the OMML element m:X
            string local = word.Substring(1);
            XElement parent = Container();
            var element = new XElement(M + local);
            parent.Add(element);
            frame.Element = element;
            if (local == "r") { currentRun = element; runFrame = frame; ApplyColor(element); }
            else currentRun = null;
            if (ValueProperties.Contains(local)) frame.Value = new StringBuilder();
        }
        // the equation: m:oMathPara (display) or m:oMath (inline); a bare argument list is wrapped as one equation
        XElement? zone = root.Descendants().FirstOrDefault(e => e.Name.LocalName is "oMathPara" or "oMath");
        if (zone is null) return null;
        return zone;

        static int HexValue(byte c) => c is >= (byte)'0' and <= (byte)'9' ? c - '0' : c is >= (byte)'a' and <= (byte)'f' ? c - 'a' + 10 : c is >= (byte)'A' and <= (byte)'F' ? c - 'A' + 10 : -100;
        static bool IsLetter(byte c) => c is >= (byte)'a' and <= (byte)'z' or >= (byte)'A' and <= (byte)'Z';
    }

    private static string Decode(ReadOnlySpan<byte> bytes, int codePage)
    {
        var chars = new char[bytes.Length * 2 + 4];
        int count = RtfCodePages.Decode(bytes, codePage, chars);
        return new string(chars, 0, count);
    }
}

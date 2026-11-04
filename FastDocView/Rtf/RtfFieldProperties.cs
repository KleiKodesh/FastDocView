using System.Globalization;
using System.Text;

namespace FastDocView.Rtf;

/// <summary>
/// The document properties .rtf fields show, from the header's {\info} group (RTF 1.5 spec "Information Group": \title,
/// \subject, \author, \manager, \company, \operator, \category, \keywords, \doccomm, \version, \edmins, \nofwords,
/// \nofchars, {\creatim / \revtim / \printim \yr \mo \dy \hr \min \sec}), the user-defined properties
/// ({\*\userprops {\propname name}\proptypeN{\staticval value}}, same section) and document variables
/// ({\*\docvar {name}{value}}: LibreOffice rtfdocumentimpl.cxx:3738-3759 - the first inner group is the name, the next the
/// value). Read once, on first use, by a small forward scan of the file's start that stops at the body's first paragraph
/// (the info group is part of the header) - no document text is held.
/// </summary>
internal sealed class RtfFieldProperties
{
    private readonly Dictionary<string, string> _properties = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, DateTime> _dates = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _variables = new(StringComparer.Ordinal);

    public string? Property(string name) => _properties.TryGetValue(name, out string? value) ? value : null;
    public DateTime? Date(string name) => _dates.TryGetValue(name, out DateTime value) ? value : null;
    public string? Variable(string name) => _variables.TryGetValue(name, out string? value) ? value : null;

    public RtfFieldProperties(Stream file)
    {
        file.Position = 0;
        var reader = new BufferedStream(file, 64 * 1024);
        Encoding encoding = Markup.MarkupTextEncoding.FromCodePage(1252);
        int depth = 0, unicodeSkip = 1, pendingSkip = 0;
        // the destination of each open group: what its text is for
        var groups = new Stack<(string destination, StringBuilder text)>();
        string? propertyName = null, docVariableName = null;
        var time = new Dictionary<string, int>();
        var bytes = new List<byte>();
        StringBuilder? Current() => groups.Count > 0 ? groups.Peek().text : null;
        int Part(string name, int fallback) => time.TryGetValue(name, out int value) ? value : fallback;
        void FlushBytes()
        {
            if (bytes.Count == 0) return;
            Current()?.Append(encoding.GetString(bytes.ToArray()));
            bytes.Clear();
        }
        while (true)
        {
            int b = reader.ReadByte();
            if (b < 0) break;
            if (b == '{') { FlushBytes(); depth++; groups.Push(("", new StringBuilder())); continue; }
            if (b == '}')
            {
                FlushBytes();
                depth--;
                if (groups.Count == 0) break;
                (string destination, StringBuilder text) = groups.Pop();
                string value = text.ToString().Trim();
                string parent = groups.Count > 0 ? groups.Peek().destination : "";
                switch (destination)
                {
                    case "title": _properties["TITLE"] = value; break;
                    case "subject": _properties["SUBJECT"] = value; break;
                    case "author": _properties["AUTHOR"] = value; break;
                    case "manager": _properties["MANAGER"] = value; break;
                    case "company": _properties["COMPANY"] = value; break;
                    case "operator": _properties["LASTSAVEDBY"] = value; break;
                    case "category": _properties["CATEGORY"] = value; break;
                    case "keywords": _properties["KEYWORDS"] = value; break;
                    case "doccomm": _properties["COMMENTS"] = value; break;
                    case "creatim" or "revtim" or "printim":
                        if (time.TryGetValue("yr", out int year) && year > 0)
                        {
                            try
                            {
                                var date = new DateTime(year, Math.Max(1, Part("mo", 1)), Math.Max(1, Part("dy", 1)),
                                                        Part("hr", 0), Part("min", 0), Part("sec", 0));
                                _dates[destination == "creatim" ? "CREATEDATE" : destination == "revtim" ? "SAVEDATE" : "PRINTDATE"] = date;
                            }
                            catch (ArgumentOutOfRangeException) { }
                        }
                        time.Clear();
                        break;
                    case "propname": propertyName = value; break;
                    case "staticval": if (propertyName is not null) _properties[propertyName] = value; propertyName = null; break;
                    case "" when parent == "docvar":
                        if (docVariableName is null) docVariableName = value;
                        else { _variables[docVariableName] = value; docVariableName = null; }
                        break;
                }
                if (depth <= 0) break;
                continue;
            }
            if (b == '\\')
            {
                int next = reader.ReadByte();
                if (next < 0) break;
                if (next == '\'')
                {
                    int high = reader.ReadByte(), low = reader.ReadByte();
                    if (pendingSkip > 0) { pendingSkip--; continue; }
                    if (int.TryParse(((char)high).ToString() + (char)low, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int code)) bytes.Add((byte)code);
                    continue;
                }
                if (!char.IsLetter((char)next))
                {
                    if (next is '\\' or '{' or '}') { if (pendingSkip > 0) pendingSkip--; else bytes.Add((byte)next); }
                    continue;
                }
                var word = new StringBuilder().Append((char)next);
                int c;
                while ((c = reader.ReadByte()) >= 0 && char.IsLetter((char)c)) word.Append((char)c);
                var number = new StringBuilder();
                if (c == '-') { number.Append('-'); c = reader.ReadByte(); }
                while (c >= 0 && char.IsDigit((char)c)) { number.Append((char)c); c = reader.ReadByte(); }
                int parameter = int.TryParse(number.ToString(), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out int p) ? p : 0;
                bool hasParameter = number.Length > 0;
                string name = word.ToString();
                // the delimiter: a space is part of the word; anything else is the next token
                if (c >= 0 && c != ' ') reader.Seek(-1, SeekOrigin.Current);
                FlushBytes();
                switch (name)
                {
                    case "ansicpg" when hasParameter:
                        try { encoding = Markup.MarkupTextEncoding.FromCodePage(parameter); } catch (Exception exception) when (exception is ArgumentException or NotSupportedException) { }
                        break;
                    case "uc": unicodeSkip = Math.Max(0, parameter); break;
                    case "u" when hasParameter:
                        Current()?.Append((char)(parameter < 0 ? parameter + 65536 : parameter));
                        pendingSkip = unicodeSkip;
                        break;
                    case "title" or "subject" or "author" or "manager" or "company" or "operator" or "category" or "keywords" or "doccomm"
                         or "creatim" or "revtim" or "printim" or "propname" or "staticval" or "docvar" or "info" or "userprops":
                        if (groups.Count > 0) groups.Push((name, groups.Pop().text));
                        break;
                    case "yr" or "mo" or "dy" or "hr" or "min" or "sec": time[name] = parameter; break;
                    case "version": _properties["REVNUM"] = parameter.ToString(CultureInfo.InvariantCulture); break;
                    case "edmins": _properties["EDITTIME"] = parameter.ToString(CultureInfo.InvariantCulture); break;
                    case "nofwords": _properties["NUMWORDS"] = parameter.ToString(CultureInfo.InvariantCulture); break;
                    case "nofchars": _properties["NUMCHARS"] = parameter.ToString(CultureInfo.InvariantCulture); break;
                    // the body has begun: the header (and its info group) is over
                    case "par" or "pard" or "sectd" or "plain" when depth == 1: return;
                    case "bin" when hasParameter: reader.Seek(parameter, SeekOrigin.Current); break;
                }
                continue;
            }
            if (b is '\r' or '\n') continue;
            if (pendingSkip > 0) { pendingSkip--; continue; }
            bytes.Add((byte)b);
        }
    }
}

using System.Globalization;
using System.Text;
using FastDocView.Html;

namespace FastDocView.Fields;

/// <summary>
/// The EQ field (ECMA-376 Part 1 17.16.5.20) as MathML Core (Chromium / WebView2 draw it natively): \a array (\al \ac \ar
/// \coN \vsN \hsN), \b brackets (\lc\X \rc\X \bc\X), \d displacement (\foN \baN \li), \f fraction, \i integral / sum /
/// product (\su \pr \in \fc\X \vc\X), \l list, \o overstrike (\al \ac \ar), \r radical, \s super / subscript (\upN \doN
/// \aiN \diN), \x box (\to \bo \le \ri). Elements are separated by "," (Word also accepts ";" where the list separator is
/// ";"); \, \( \) \\ escape a character. MathML Core has no menclose and no negative mspace: \x is a bordered mrow (CSS)
/// and \ba (move back) a negative margin - the nearest Core equivalents.
/// </summary>
public static class WordFieldEquation
{
    public static string ToHtml(string code)
    {
        var parser = new Parser(code);
        string body = parser.Sequence(stopAtSeparator: false);
        return "<math class=\"field-equation\" display=\"inline\"><mrow>" + body + "</mrow></math>";
    }

    private sealed class Parser(string code)
    {
        private int _position;

        /// <summary>An open sequence: the root, a bare "( ... )" group or one element of a switch's list.</summary>
        private sealed class SequenceFrame(SequenceKind kind)
        {
            public readonly SequenceKind Kind = kind;
            public readonly StringBuilder Html = new();
            public readonly StringBuilder Text = new();
            public bool StopAtSeparator => Kind == SequenceKind.Element;

            public void FlushText()
            {
                if (Text.Length == 0) return;
                Html.Append("<mtext>").Append(WordHtmlText.Encode(Text.ToString())).Append("</mtext>");
                Text.Clear();
            }
        }

        private enum SequenceKind { Root, Group, Element }

        /// <summary>A switch whose element list is being read.</summary>
        private sealed record SwitchFrame(string Name, List<(string name, string value)> Options, List<string> Elements);

        /// <summary>
        /// Text and switches up to the end. Nested groups and element lists are walked with an explicit stack (a deeply
        /// nested EQ code must not overflow the thread's stack; Word 16 itself hangs on 300 nested \f, so it defines no limit).
        /// </summary>
        public string Sequence(bool stopAtSeparator)
        {
            var sequences = new Stack<SequenceFrame>();
            var switches = new Stack<SwitchFrame>();
            sequences.Push(new SequenceFrame(stopAtSeparator ? SequenceKind.Element : SequenceKind.Root));
            while (true)
            {
                SequenceFrame current = sequences.Peek();
                if (_position >= code.Length || code[_position] == ')' || current.StopAtSeparator && code[_position] is ',' or ';')
                {
                    current.FlushText();
                    string body = current.Html.ToString();
                    if (sequences.Count == 1) return body;
                    sequences.Pop();
                    if (current.Kind == SequenceKind.Group)
                    {
                        StringBuilder parent = sequences.Peek().Html;
                        parent.Append(body);
                        if (_position < code.Length && code[_position] == ')') { _position++; parent.Append("<mo>)</mo>"); }
                        continue;
                    }
                    SwitchFrame open = switches.Peek();
                    open.Elements.Add(body);
                    if (_position < code.Length && code[_position++] != ')') { sequences.Push(new SequenceFrame(SequenceKind.Element)); continue; }
                    switches.Pop();
                    sequences.Peek().Html.Append(Render(open.Name, open.Options, open.Elements));
                    continue;
                }
                char c = code[_position];
                if (c == '\\' && _position + 1 < code.Length)
                {
                    char next = code[_position + 1];
                    if (!char.IsLetter(next)) { current.Text.Append(next); _position += 2; continue; }
                    current.FlushText();
                    SwitchFrame started = SwitchStart();
                    // the parenthesised element list after the switch and its options
                    while (_position < code.Length && char.IsWhiteSpace(code[_position])) _position++;
                    if (_position >= code.Length || code[_position] != '(') { current.Html.Append(Render(started.Name, started.Options, started.Elements)); continue; }
                    _position++;
                    switches.Push(started);
                    sequences.Push(new SequenceFrame(SequenceKind.Element));
                    continue;
                }
                if (c == '(')
                {
                    // a bare parenthesis group is text with its parentheses
                    _position++;
                    current.FlushText();
                    current.Html.Append("<mo>(</mo>");
                    sequences.Push(new SequenceFrame(SequenceKind.Group));
                    continue;
                }
                current.Text.Append(c);
                _position++;
            }
        }

        private string ReadWord()
        {
            int start = _position;
            while (_position < code.Length && char.IsLetter(code[_position])) _position++;
            return code.Substring(start, _position - start);
        }

        private int ReadNumber()
        {
            int start = _position;
            while (_position < code.Length && char.IsDigit(code[_position])) _position++;
            return int.TryParse(code.Substring(start, _position - start), NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) ? value : 0;
        }

        /// <summary>An option's character argument (\lc\[ ): the character after the backslash, or the next character.</summary>
        private string ReadCharacter()
        {
            if (_position < code.Length && code[_position] == '\\' && _position + 1 < code.Length) { _position += 2; return code[_position - 1].ToString(); }
            if (_position < code.Length) { _position++; return code[_position - 1].ToString(); }
            return "";
        }

        /// <summary>A switch's name and options (\xxN or \xx\c before its element list).</summary>
        private SwitchFrame SwitchStart()
        {
            _position++;                                  // the backslash
            string name = ReadSwitchName();
            var options = new List<(string name, string value)>();
            while (true)
            {
                // options may be written apart from the switch and each other (\a \co2 \al(...))
                int save = _position;
                while (_position < code.Length && char.IsWhiteSpace(code[_position])) _position++;
                if (!(_position + 1 < code.Length && code[_position] == '\\' && char.IsLetter(code[_position + 1]))) { _position = save; break; }
                _position++;
                string option = ReadOptionName(name);
                if (option.Length == 0) { _position = save; break; }
                string value = option is "lc" or "rc" or "bc" or "fc" or "vc" ? ReadCharacter() : ReadNumber().ToString(CultureInfo.InvariantCulture);
                options.Add((option, value));
            }
            return new SwitchFrame(name, options, []);
        }

        private static string Render(string name, List<(string name, string value)> options, List<string> elements)
        {
            string? Option(string option) => options.Where(o => o.name == option).Select(o => o.value).Cast<string?>().FirstOrDefault();
            bool Has(string option) => options.Any(o => o.name == option);
            switch (name)
            {
                case "f":
                    return "<mfrac><mrow>" + At(elements, 0) + "</mrow><mrow>" + At(elements, 1) + "</mrow></mfrac>";
                case "r":
                    return elements.Count >= 2 ? "<mroot><mrow>" + elements[1] + "</mrow><mrow>" + elements[0] + "</mrow></mroot>" : "<msqrt><mrow>" + At(elements, 0) + "</mrow></msqrt>";
                case "b":
                {
                    string left = Option("bc") ?? Option("lc") ?? "(";
                    string right = Option("bc") is { } both ? Mirror(both) : Option("rc") ?? ")";
                    return "<mrow><mo stretchy=\"true\">" + WordHtmlText.Encode(left) + "</mo><mrow>" + string.Concat(elements) + "</mrow><mo stretchy=\"true\">"
                           + WordHtmlText.Encode(right) + "</mo></mrow>";
                }
                case "l": return "<mrow>" + string.Join("<mo>,</mo>", elements) + "</mrow>";
                case "s":
                {
                    if (elements.Count > 1)
                        return "<mtable class=\"eq-stack\">" + string.Concat(elements.Select(e => "<mtr><mtd><mrow>" + e + "</mrow></mtd></mtr>")) + "</mtable>";
                    int up = int.TryParse(Option("up"), out int u) ? u : 0, down = int.TryParse(Option("do"), out int d) ? d : 0;
                    int offset = up - down;
                    return "<mpadded voffset=\"" + offset.ToString(CultureInfo.InvariantCulture) + "pt\"><mrow>" + At(elements, 0) + "</mrow></mpadded>";
                }
                case "i":
                {
                    string symbol = Option("fc") ?? Option("vc") ?? (Has("su") ? "∑" : Has("pr") ? "∏" : "∫");
                    string op = "<mo largeop=\"true\" movablelimits=\"false\">" + WordHtmlText.Encode(symbol) + "</mo>";
                    string limits = Has("in") || symbol == "∫"
                        ? "<msubsup>" + op + "<mrow>" + At(elements, 0) + "</mrow><mrow>" + At(elements, 1) + "</mrow></msubsup>"
                        : "<munderover>" + op + "<mrow>" + At(elements, 0) + "</mrow><mrow>" + At(elements, 1) + "</mrow></munderover>";
                    return "<mrow>" + limits + "<mrow>" + At(elements, 2) + "</mrow></mrow>";
                }
                case "a":
                {
                    // more columns than elements is one row of all the elements (no empty cells; also keeps i + columns in range)
                    int columns = Math.Min(Math.Max(1, int.TryParse(Option("co"), out int co) ? co : 1), Math.Max(1, elements.Count));
                    string align = Has("al") ? "left" : Has("ar") ? "right" : "center";
                    var table = new StringBuilder("<mtable columnalign=\"" + align + "\">");
                    for (int i = 0; i < elements.Count; i += columns)
                    {
                        table.Append("<mtr>");
                        for (int j = i; j < i + columns; j++) table.Append("<mtd><mrow>").Append(j < elements.Count ? elements[j] : "").Append("</mrow></mtd>");
                        table.Append("</mtr>");
                    }
                    return table.Append("</mtable>").ToString();
                }
                case "o":
                {
                    // overstrike: every element drawn on the same place (all but the last take no width)
                    var html = new StringBuilder("<mrow>");
                    for (int i = 0; i < elements.Count; i++)
                        html.Append(i < elements.Count - 1 ? "<mpadded width=\"0\"><mrow>" + elements[i] + "</mrow></mpadded>" : "<mrow>" + elements[i] + "</mrow>");
                    return html.Append("</mrow>").ToString();
                }
                case "x":
                {
                    bool any = Has("to") || Has("bo") || Has("le") || Has("ri");
                    var css = new StringBuilder("padding:0 2px;");
                    foreach ((string option, string side) in (ReadOnlySpan<(string, string)>)[("to", "top"), ("bo", "bottom"), ("le", "left"), ("ri", "right")])
                        if (!any || Has(option)) css.Append("border-").Append(side).Append(":1px solid currentColor;");
                    return "<mrow style=\"" + css + "\">" + string.Concat(elements) + "</mrow>";
                }
                case "d":
                {
                    int forward = int.TryParse(Option("fo"), out int f) ? f : 0, back = int.TryParse(Option("ba"), out int b) ? b : 0;
                    string style = Has("li") ? " style=\"text-decoration:underline\"" : "";
                    if (back > 0) return "<mrow style=\"margin-inline-start:-" + back.ToString(CultureInfo.InvariantCulture) + "pt\"" + "></mrow>";
                    return "<mspace width=\"" + forward.ToString(CultureInfo.InvariantCulture) + "pt\"" + style + "></mspace>";
                }
                default: return string.Concat(elements);
            }
        }

        private string ReadSwitchName()
        {
            // switch names are one letter (\a \b \d \f \i \l \o \r \s \x)
            if (_position < code.Length && char.IsLetter(code[_position])) return char.ToLowerInvariant(code[_position++]).ToString();
            return "";
        }

        /// <summary>A two-letter option of the switch (\al, \co3, \lc\{ ...); "" when the next backslash is not one.</summary>
        private string ReadOptionName(string switchName)
        {
            if (_position + 1 >= code.Length) return "";
            string option = code.Substring(_position, 2).ToLowerInvariant();
            string[] known = switchName switch
            {
                "a" => ["al", "ac", "ar", "co", "vs", "hs"], "b" => ["lc", "rc", "bc"], "d" => ["fo", "ba", "li"], "i" => ["su", "pr", "in", "fc", "vc"],
                "o" => ["al", "ac", "ar"], "s" => ["ai", "di", "up", "do"], "x" => ["to", "bo", "le", "ri"], _ => [],
            };
            if (Array.IndexOf(known, option) < 0) return "";
            _position += 2;
            return option;
        }

        private static string At(List<string> elements, int index) => index < elements.Count ? elements[index] : "";

        /// <summary>\bc\X: the right bracket mirrors the left one ( ( → ), [ → ], { → }, &lt; → &gt; ).</summary>
        private static string Mirror(string bracket) => bracket switch { "(" => ")", "[" => "]", "{" => "}", "<" => ">", _ => bracket };
    }
}

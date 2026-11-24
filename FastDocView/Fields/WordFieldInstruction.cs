using System.Text;

namespace FastDocView.Fields;

/// <summary>One word of a field instruction: a switch (<c>\x</c>, upper-cased as LibreOffice does), or text (quotes removed).</summary>
public readonly record struct WordFieldToken(string Text, bool IsSwitch, bool IsQuoted);

/// <summary>
/// A Word field instruction ("REF bm \h \* MERGEFORMAT") split into its name, plain arguments and switches.
/// Tokenizing follows LibreOffice's .docx importer (sw/source/writerfilter/dmapper/DomainMapper_Impl.cxx:5918-6015
/// lcl_ExtractToken / splitFieldCommand): whitespace separates words outside quotes; inside quotes a backslash takes the next
/// character literally; outside quotes <c>\\</c> is one backslash and <c>\x</c> is a switch; a leading <c>=</c> is the formula
/// field. Which switches take an argument is field specific (the field descriptions of ECMA-376 Part 1 17.16.5); the
/// general switches <c>\*</c>, <c>\#</c> and <c>\@</c> always do (17.16.4).
/// </summary>
public sealed class WordFieldInstruction
{
    public string Code { get; }
    /// <summary>The field name upper-cased ("REF", "SEQ"); "=" for a formula; "" for an empty instruction.</summary>
    public string Name { get; }
    public IReadOnlyList<WordFieldToken> Tokens { get; }
    /// <summary>The words that are neither switches nor a switch's argument, in order.</summary>
    public IReadOnlyList<string> Arguments { get; }
    private readonly List<(string name, string? argument)> _switches;
    public IReadOnlyList<(string Name, string? Argument)> Switches => _switches.Select(s => (s.name, s.argument)).ToList();

    private WordFieldInstruction(string code, string name, List<WordFieldToken> tokens, List<string> arguments, List<(string, string?)> switches)
    {
        Code = code; Name = name; Tokens = tokens; Arguments = arguments; _switches = switches;
    }

    /// <summary>Whether the switch (e.g. <c>\h</c>, any case) is present.</summary>
    public bool Has(string switchName)
    {
        foreach ((string name, _) in _switches) if (name.Equals(switchName, StringComparison.OrdinalIgnoreCase)) return true;
        return false;
    }

    /// <summary>The argument of the first such switch, "" when it has none, null when absent.</summary>
    public string? SwitchArgument(string switchName)
    {
        foreach ((string name, string? argument) in _switches) if (name.Equals(switchName, StringComparison.OrdinalIgnoreCase)) return argument ?? "";
        return null;
    }

    /// <summary>Every <c>\*</c> argument, in order (several may be given: a number format and a case format).</summary>
    public IEnumerable<string> GeneralFormats
    {
        get { foreach ((string name, string? argument) in _switches) if (name == "\\*" && argument is not null) yield return argument; }
    }

    public string? NumericPicture => SwitchArgument("\\#") is { Length: > 0 } picture ? picture : null;
    public string? DatePicture => SwitchArgument("\\@") is { Length: > 0 } picture ? picture : null;

    /// <summary>The expression of a formula field: the code after "=" up to its first switch outside quotes.</summary>
    public string FormulaExpression
    {
        get
        {
            int start = Code.IndexOf('=');
            if (start < 0) return "";
            bool quoted = false;
            for (int i = start + 1; i < Code.Length; i++)
            {
                if (Code[i] == '"') quoted = !quoted;
                else if (Code[i] == '\\' && !quoted && i + 1 < Code.Length && Code[i + 1] is '*' or '#' or '@') return Code.Substring(start + 1, i - start - 1);
            }
            return Code.Substring(start + 1);
        }
    }

    /// <summary>
    /// Switches that take an argument, by field (ECMA-376 Part 1 17.16.5, each field's "Field-specific switches"); the
    /// general switches always take one.
    /// </summary>
    private static readonly Dictionary<string, string[]> ArgumentSwitches = new(StringComparer.Ordinal)
    {
        ["HYPERLINK"] = ["\\L", "\\O", "\\T"],
        ["SEQ"] = ["\\R", "\\S"],
        ["TOC"] = ["\\A", "\\B", "\\C", "\\D", "\\F", "\\L", "\\O", "\\P", "\\S", "\\T"],
        ["INDEX"] = ["\\B", "\\C", "\\D", "\\E", "\\F", "\\G", "\\H", "\\K", "\\L", "\\P", "\\S"],
        ["TOA"] = ["\\B", "\\C", "\\D", "\\E", "\\G", "\\L", "\\S"],
        ["REF"] = ["\\D"],
        ["NOTEREF"] = [],
        ["PAGEREF"] = [],
        ["LISTNUM"] = ["\\L", "\\S"],
        ["XE"] = ["\\F", "\\R", "\\T"],
        ["TA"] = ["\\C", "\\L", "\\R", "\\S"],
        ["TC"] = ["\\F", "\\L"],
        ["RD"] = [],
        ["SYMBOL"] = ["\\F", "\\S"],
        ["ADVANCE"] = ["\\D", "\\U", "\\L", "\\R", "\\X", "\\Y"],
        ["INCLUDETEXT"] = ["\\C"],
        ["INCLUDEPICTURE"] = ["\\C"],
        ["MERGEFIELD"] = ["\\B", "\\F"],
        ["AUTONUM"] = ["\\S"],
        ["AUTONUMLGL"] = ["\\S"],
        ["BARCODE"] = ["\\B", "\\F", "\\U"],
        ["DOCPROPERTY"] = [],
        ["SET"] = [],
    };

    /// <summary>The name of a field Word gives no type (Field.Type -1): an empty code or a keyword a note reference broke. Never a token (a space).</summary>
    public const string NoTypeName = " ";

    public static WordFieldInstruction Parse(string code)
    {
        if (string.IsNullOrWhiteSpace(code)) return new WordFieldInstruction(code, NoTypeName, [], [], []);
        // a note reference inside the code (WordFieldRuby.NoteMarkPlaceholder) is only drawn by an EQ ruby; for any other field it
        // separates tokens as the space it was read as before (its mark is drawn after the field's result)
        if (code.IndexOf(WordFieldRuby.NoteMarkPlaceholder) >= 0)
        {
            // Word 16 (2026-11-24, test-data/word-ruby/note-in-field-keyword-v1.rtf / .docx / .doc): a note reference before the
            // field's keyword or touching its end ("MERGEFIELD{note} name", "{note}MERGEFIELD") leaves the field without a type
            // (Field.Type -1): it shows nothing, locked or not, and no mark. "=" is a keyword by itself ("={note} 1+2" computes).
            if (IsKeywordBrokenByNote(code)) return new WordFieldInstruction(code.Replace(WordFieldRuby.NoteMarkPlaceholder.ToString(), ""), NoTypeName, [], [], []);
            if (!code.TrimStart().StartsWith("EQ", StringComparison.OrdinalIgnoreCase)) code = code.Replace(WordFieldRuby.NoteMarkPlaceholder, ' ');
        }
        var tokens = new List<WordFieldToken>();
        int index = 0;
        // DomainMapper_Impl.cxx:6024-6031: a field may start with a stray backslash ("\REF ..."); Word skips it
        if (code.Length >= 2 && code.TrimStart() is { Length: >= 2 } trimmed && trimmed[0] == '\\' && trimmed[1] != '\\' && !char.IsWhiteSpace(trimmed[1]))
            index = code.IndexOf('\\') + 1;
        string name = "";
        while (index < code.Length)
        {
            if (ExtractToken(code, ref index, out bool isSwitch, out bool isQuoted) is not { } token) continue;
            if (name.Length == 0 && tokens.Count == 0 && !isSwitch) { name = token == "=" ? "=" : token.ToUpperInvariant(); continue; }
            tokens.Add(new WordFieldToken(token, isSwitch, isQuoted));
        }
        ArgumentSwitches.TryGetValue(name, out string[]? takesArgument);
        var arguments = new List<string>();
        var switches = new List<(string, string?)>();
        for (int i = 0; i < tokens.Count; i++)
        {
            WordFieldToken token = tokens[i];
            if (!token.IsSwitch) { arguments.Add(token.Text); continue; }
            string upper = token.Text.ToUpperInvariant();
            bool hasArgument = upper is "\\*" or "\\#" or "\\@" || takesArgument is not null && Array.IndexOf(takesArgument, upper) >= 0;
            if (hasArgument && i + 1 < tokens.Count && !tokens[i + 1].IsSwitch) { switches.Add((token.Text, tokens[i + 1].Text)); i++; }
            else switches.Add((token.Text, null));
        }
        return new WordFieldInstruction(code, name, tokens, arguments, switches);
    }

    /// <summary>Whether a note reference (its placeholder) comes before the code's first token or right after it, with no space between.</summary>
    private static bool IsKeywordBrokenByNote(string code)
    {
        int index = 0;
        while (index < code.Length && char.IsWhiteSpace(code[index])) index++;
        if (index == code.Length) return false;
        if (code[index] == '=') return false;
        for (; index < code.Length && !char.IsWhiteSpace(code[index]); index++)
            if (code[index] == WordFieldRuby.NoteMarkPlaceholder) return true;
        return false;
    }

    /// <summary>DomainMapper_Impl.cxx:5918 lcl_ExtractToken, with any whitespace as a separator.</summary>
    private static string? ExtractToken(string code, ref int index, out bool isSwitch, out bool isQuoted)
    {
        isSwitch = false;
        isQuoted = false;
        var token = new StringBuilder();
        bool quoted = false;
        for (; index < code.Length; index++)
        {
            char c = code[index];
            switch (c)
            {
                case '\\':
                    if (index == code.Length - 1) { index++; return token.Length > 0 ? token.ToString() : null; }
                    char next = code[index + 1];
                    if (quoted || next == '\\') { index++; token.Append(next); break; }
                    if (token.Length == 0) { isSwitch = true; index += 2; return code.Substring(index - 2, 2); }
                    return token.ToString();                                  // read the switch again next time
                case '"':
                    if (quoted) { index++; isQuoted = true; return token.ToString(); }
                    if (token.Length > 0) return token.ToString();
                    quoted = true;
                    break;
                case '=' when token.Length == 0 && !quoted:
                    index++;
                    return "=";
                default:
                    if (char.IsWhiteSpace(c) && !quoted)
                    {
                        if (token.Length > 0) { index++; return token.ToString(); }
                        break;
                    }
                    token.Append(c);
                    break;
            }
        }
        isQuoted = quoted;
        return token.Length > 0 || quoted ? token.ToString() : null;
    }
}

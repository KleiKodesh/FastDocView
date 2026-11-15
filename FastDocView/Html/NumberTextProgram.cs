using System.Text;
using System.Text.RegularExpressions;

namespace FastDocView.Html;

/// <summary>
/// One compiled Soros program (libnumbertext src/Soros.cxx, BSD / LGPL dual licence; data in Html/NumberText): the number-to-text
/// rules LibreOffice's NATNUM12 runs for spelled-out list numbers (i18npool nativenumbersupplier.cxx getNumberText). A line-for-line
/// port: the program text is rewritten into regex / replacement pairs, a run tries the patterns in order on the whole input and
/// expands calls "$(...)" recursively. The replacement strings keep std::regex's ECMAScript format ($$, $&amp;, $n / $nn), expanded here.
/// </summary>
internal sealed class NumberTextProgram
{
    private const int IterationLimit = 250, CallLimit = 10000;
    private const string Separator = "\uE00A";
    private const string Meta = "\\\";#", Meta2 = "$()|[]";
    private const string Codes = "\uE000\uE001\uE002\uE003", Codes2 = "\uE004\uE005\uE006\uE007\uE008\uE009";
    private const char Slash = '\uE000', Pipe = '\uE003';

    // \$, \(, \), \| -> U+E000..U+E003 (Soros.cxx func)
    private static readonly Regex Function = new(Translate(
        "(?:\\|?(?:\\$\\()+)?" + "(\\|?\\$\\(([^\\(\\)]*)\\)\\|?)" + "(?:\uE00A?\\)+\\|?)?", Meta2.Substring(0, 4), Codes, "\\"), RegexOptions.CultureInvariant);

    private readonly List<Regex> _patterns = [];
    private readonly List<string> _values = [];
    private readonly List<bool> _begins = [], _ends = [];

    public NumberTextProgram(string program, string language)
    {
        program = program.Replace("\r", "");
        program = Translate(program, Meta, Codes, "\\");
        // switch off all country-dependent lines, and switch on the requested ones
        program = Regex.Replace(program, "(^|[\n;])([^\n;#]*#[^\n]*\\[:[^\n:\\]]*:\\][^\n]*)", "$1#$2");
        string filtered = language.Replace("_", "-");
        program = Regex.Replace(program, "(^|[\n;])#([^\n;#]*#[^\n]*\\[:" + filtered + ":\\][^\n]*)", "$1$2");
        program = Regex.Replace(program, "(#[^\n]*)?(\n|\\z)", ";");
        if (program.IndexOf("__numbertext__", StringComparison.Ordinal) < 0) program = "__numbertext__;" + program;
        program = program.Replace("__numbertext__",
            "\"([a-z][-a-z]* )?0+(0|[1-9]\\d*)\" $(\\1\\2);"
            + Separator + "(.*)" + Separator + "(.+)" + Separator + "(.*) \\1\\2\\3;"
            + Separator + ".*" + Separator + Separator + ".*");

        var line = new Regex("^\\s*(\"[^\"]*\"|[^\\s]*)\\s*(.*[^\\s])?\\s*$");
        var macro = new Regex("^(?:== *([^ ]*) *==)$");
        string currentMacro = "";
        int position = 0, oldPosition = 0;
        while ((position = program.IndexOf(';', position)) >= 0)
        {
            string original = program.Substring(oldPosition, position - oldPosition), text = original;
            Match parts;
            if (currentMacro.Length > 0 && original.Length > 0 && (parts = line.Match(original)).Success)
            {
                string s = StripQuotes(parts.Groups[1].Value);
                string space = s.Length == 0 ? "" : " ";
                text = s.Length > 0 && s[0] == '^'
                    ? "\"^" + currentMacro + space + s.Substring(1) + "\" " + parts.Groups[2].Value
                    : "\"" + currentMacro + space + s + "\" " + parts.Groups[2].Value;
            }
            Match macroMatch;
            if (original.Length > 0 && (macroMatch = macro.Match(original)).Success) currentMacro = macroMatch.Groups[1].Value;
            else if (text.Length > 0 && (parts = line.Match(text)).Success)
            {
                string s = StripQuotes(parts.Groups[1].Value);
                s = Translate(s, Codes.Substring(1), Meta.Substring(1), "");
                s = s.Replace(Slash.ToString(), "\\\\");
                _begins.Add(s.Length > 0 && s[0] == '^');
                _ends.Add(s.Length > 0 && s[s.Length - 1] == '$');
                s = "^" + (s.StartsWith("^", StringComparison.Ordinal) ? s.Substring(1) : s);
                s = (s.EndsWith("$", StringComparison.Ordinal) ? s.Substring(0, s.Length - 1) : s) + "$";
                Regex pattern;
                try { pattern = Compile(s); }
                catch (ArgumentException) { break; }      // Soros.cxx: a bad regex ends the program
                _patterns.Add(pattern);
                string s2 = parts.Groups[2].Success ? StripQuotes(parts.Groups[2].Value) : "";
                s2 = Translate(s2, Meta2, Codes2, "\\");
                s2 = Regex.Replace(s2, "^\\[[$](\\d\\d?|\\([^\\)]+\\))", "$$(" + Separator + Separator + "|$$$1" + Separator);
                s2 = Regex.Replace(s2, "\\[([^$\\[\\\\]*)[$](\\d\\d?|\\([^\\)]+\\))", "$$(" + Separator + "$1" + Separator + "$$$2" + Separator);
                s2 = s2.Replace("]", ")");
                s2 = Regex.Replace(s2, "([$]\\d|\\))\\|[$]", "$1||$$");
                s2 = Translate(s2, Codes, Meta, "");
                s2 = Translate(s2, Meta2.Substring(0, 4), Codes, "");
                s2 = Translate(s2, Codes2, Meta2, "");
                s2 = s2.Replace("$", "\\$");
                s2 = Regex.Replace(s2, "\uE000(\\d)", "\uE000\uE001$$$1\uE002");
                s2 = Regex.Replace(s2, "\\\\([1-9])", "$$0$1");
                s2 = Regex.Replace(s2, "\\\\0", "$$0");
                s2 = Regex.Replace(s2, "\\\\n", "\n");
                _values.Add(s2);
            }
            position++;
            oldPosition = position;
        }
    }

    /// <summary>
    /// A whole-input match as std::regex_match: the pattern wrapped in \A(?: )\z so backtracking looks for a full match. ECMAScript
    /// classes (\d = ASCII digits) when .NET's ECMAScript mode accepts the pattern.
    /// </summary>
    private static Regex Compile(string pattern)
    {
        string whole = "\\A(?:" + pattern + ")\\z";
        try { return new Regex(whole, RegexOptions.ECMAScript | RegexOptions.CultureInvariant); }
        catch (ArgumentException) { return new Regex(whole, RegexOptions.CultureInvariant); }
    }

    private static string StripQuotes(string s)
    {
        if (s.StartsWith("\"", StringComparison.Ordinal)) s = s.Substring(1);
        if (s.EndsWith("\"", StringComparison.Ordinal)) s = s.Substring(0, s.Length - 1);
        return s;
    }

    private static string Translate(string s, string chars, string chars2, string delimiter)
    {
        for (int i = 0; i < chars.Length; i++) s = s.Replace(delimiter + chars[i], chars2[i].ToString());
        return s;
    }

    /// <summary>The program's text for <paramref name="input"/>; "" when no rule matches or a limit is hit (Soros.cxx run).</summary>
    public string Run(string input)
    {
        int level = 0, calls = 0;
        return Run(input, ref level, ref calls, true, true);
    }

    private string Run(string input, ref int level, ref int calls, bool begin, bool end)
    {
        if (level == -1) return "";
        if (++level > IterationLimit || ++calls > CallLimit) { level = -1; return ""; }
        for (int i = 0; i < _patterns.Count; i++)
        {
            if ((!begin && _begins[i]) || (!end && _ends[i])) continue;
            Match match = _patterns[i].Match(input);
            if (!match.Success) continue;
            input = Expand(_values[i], match);
            Match call;
            while ((call = Function.Match(input)).Success)
            {
                bool b = false, e = false;
                string inner = call.Groups[1].Value, whole = call.Value;
                if (inner[0] == Pipe || whole[0] == Pipe) b = true;
                else if (call.Index == 0) b = begin;
                if (inner[inner.Length - 1] == Pipe || whole[whole.Length - 1] == Pipe) e = true;
                else if (call.Index + call.Length == input.Length) e = end;
                string piece = Run(call.Groups[2].Value, ref level, ref calls, b, e);
                if (level == -1) return "";
                input = input.Substring(0, call.Groups[1].Index) + piece + input.Substring(call.Groups[1].Index + call.Groups[1].Length);
            }
            level--;
            return input;
        }
        level--;
        return "";
    }

    /// <summary>std::regex format_default (ECMAScript): $$, $&amp;, $`, $', $n and $nn (two digits taken when present; an absent group is empty).</summary>
    private static string Expand(string format, Match match)
    {
        var result = new StringBuilder(format.Length + 16);
        for (int i = 0; i < format.Length; i++)
        {
            char c = format[i];
            if (c != '$' || i + 1 >= format.Length) { result.Append(c); continue; }
            char next = format[i + 1];
            if (next == '$') { result.Append('$'); i++; }
            else if (next == '&') { result.Append(match.Value); i++; }
            else if (next is '`' or '\'') i++;      // the match is the whole input: no prefix / suffix
            else if (next is >= '0' and <= '9')
            {
                int number = next - '0';
                i++;
                if (i + 1 < format.Length && format[i + 1] is >= '0' and <= '9') { number = number * 10 + (format[i + 1] - '0'); i++; }
                if (number < match.Groups.Count && match.Groups[number].Success) result.Append(match.Groups[number].Value);
            }
            else result.Append(c);
        }
        return result.ToString();
    }
}

/// <summary>
/// libnumbertext's Numbertext (src/Numbertext.cxx) over the embedded data: the module for a language code ("he-IL": he_IL.sor,
/// else without the country he.sor; Norwegian nb / nn: no.sor), cached; run on "[function ]number". As LibreOffice's NumberText
/// service (lingucomponent numbertext.cxx:108) the input comes back unchanged when no module exists.
/// </summary>
internal static class NumberTextModules
{
    private static readonly Dictionary<string, NumberTextProgram?> Modules = new(StringComparer.Ordinal);
    private static readonly object Lock = new();

    /// <summary>The text for <paramref name="input"/> in <paramref name="language"/> (language[-Script][-Country]); null when no module exists.</summary>
    public static string? Run(string input, string language)
    {
        NumberTextProgram? program;
        lock (Lock)
        {
            if (!Modules.TryGetValue(language, out program))
            {
                program = Load(language);
                Modules[language] = program;
            }
        }
        if (program is null) return null;
        lock (program) return program.Run(input);
    }

    private static NumberTextProgram? Load(string language)
    {
        string name = language.Replace('-', '_');
        string? text = Read(name);
        if (text is null)
        {
            name = Regex.Replace(name, "[-_]..$", "");
            text = Read(name);
            if (text is null && Regex.IsMatch(language, "^n[bn]([-_]NO)?$")) text = Read("no");
        }
        return text is null ? null : new NumberTextProgram(text, language);
    }

    private static string? Read(string name)
    {
        using Stream? stream = typeof(NumberTextModules).Assembly.GetManifestResourceStream("FastDocView.NumberText." + name + ".sor");
        if (stream is null) return null;
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return reader.ReadToEnd();
    }
}

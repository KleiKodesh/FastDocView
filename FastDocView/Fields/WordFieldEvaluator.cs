using System.Globalization;
using System.Text;
using FastDocView.Html;

namespace FastDocView.Fields;

/// <summary>How a field shows in a reading view.</summary>
public enum WordFieldDisplay : byte
{
    /// <summary>Its result (cached, or computed when Word saved none / it is dirty and not locked).</summary>
    Result,
    /// <summary>A page-number field: a neutral placeholder instead of the result (flow view has no pages - NOTES.md "Fields").</summary>
    PagePlaceholder,
    /// <summary>A marker or private data (XE, TC, TA, RD, SET, ASK, PRIVATE ...): nothing shows.</summary>
    Hidden,
}

/// <summary>A form field's state (.docx w:ffData, .doc FFData, .rtf \formfield): checkbox and drop-down state, text default.</summary>
public sealed class WordFormFieldState
{
    public bool? IsChecked { get; init; }
    public IReadOnlyList<string>? Entries { get; init; }
    public int SelectedEntry { get; init; }
    public string? DefaultText { get; init; }
}

/// <summary>What a field contributes: text (searched like any shown text) or HTML (EQ, SYMBOL, form controls, placeholders).</summary>
public readonly record struct WordFieldOutput(string? Text, string? Html, bool IsSearchable)
{
    public static readonly WordFieldOutput None = new(null, null, false);
    public bool IsEmpty => Text is null && Html is null;
}

/// <summary>
/// Field semantics shared by the .docx, .doc and .rtf readers: which fields show their result, which are page placeholders or
/// hidden markers, and the result of a field Word saved without one (or marked dirty). Field names and numbers as in
/// [MS-DOC] 2.9.90 flt and LibreOffice sw/source/filter/ww8/fields.cxx:33-131; field behaviour from ECMA-376 Part 1 17.16.5 and
/// Word 16 itself (the values in WordFieldFormat / WordFieldExpression / WordFieldSequenceCounters).
/// </summary>
public static class WordFieldEvaluator
{
    /// <summary>Word 16's text when REF names a bookmark that does not exist (Hebrew UI, as written on the dev machine).</summary>
    public const string ReferenceNotFound = "שגיאה! מקור ההפניה לא נמצא.";

    public static WordFieldDisplay DisplayOf(string fieldName) => fieldName switch
    {
        // page numbers: placeholders until the paged view (ECMA-376 17.16.5.39 PAGE, .40 PAGEREF, .44 NUMPAGES, .53 SECTION, .54 SECTIONPAGES)
        "PAGE" or "PAGEREF" or "NUMPAGES" or "SECTION" or "SECTIONPAGES" => WordFieldDisplay.PagePlaceholder,
        // index / authority / contents entries, document data, mail-merge control and prompts with no visible result
        "XE" or "TC" or "TA" or "RD" or "SET" or "ASK" or "PRIVATE" or "NEXT" or "NEXTIF" or "SKIPIF" or "PRINT" => WordFieldDisplay.Hidden,
        _ => WordFieldDisplay.Result,
    };

    /// <summary>
    /// Fields Word draws from their code (they have no result: Word 16's Range.Text of a field result is empty for EQ,
    /// MACROBUTTON, GOTOBUTTON, LISTNUM, AUTONUM*, ADVANCE, FORMCHECKBOX, FORMDROPDOWN - checked 2026-11-04).
    /// </summary>
    public static bool IsDrawnFromCode(string fieldName) => fieldName is "EQ" or "MACROBUTTON" or "GOTOBUTTON" or "LISTNUM" or "AUTONUM" or "AUTONUMLGL"
        or "AUTONUMOUT" or "ADVANCE" or "FORMCHECKBOX" or "FORMDROPDOWN" or "SYMBOL";

    /// <summary>Fields counted wherever they are (SEQ, AUTONUM*, LISTNUM): the counters advance even when the result is cached.</summary>
    public static bool IsCounted(string fieldName) => fieldName is "SEQ" or "AUTONUM" or "AUTONUMLGL" or "AUTONUMOUT" or "LISTNUM";

    /// <summary>Advances the counters for a counted field; returns the value it shows (null when it is not counted).</summary>
    public static int? Count(WordFieldInstruction instruction, WordFieldSequenceCounters counters)
    {
        switch (instruction.Name)
        {
            case "SEQ": return counters.Advance(instruction);
            case "AUTONUM": case "AUTONUMLGL": case "AUTONUMOUT": return counters.AdvanceAutoNumber();
            case "LISTNUM":
            {
                int level = instruction.SwitchArgument("\\l") is { Length: > 0 } l && int.TryParse(l, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) ? Math.Max(1, Math.Min(9, parsed)) : 1;
                int? start = instruction.SwitchArgument("\\s") is { Length: > 0 } s && int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out int st) ? st : null;
                return counters.AdvanceListNumber(level, start);
            }
            default: return null;
        }
    }

    /// <summary>
    /// The page placeholder: "#" in a dotted frame, not searched, titled with the field; a PAGEREF \h still links to its
    /// bookmark (the caller wraps it).
    /// </summary>
    public static string PlaceholderHtml(WordFieldInstruction instruction)
        => "<span class=\"field-page-placeholder\" title=\"" + WordHtmlText.Encode(instruction.Name) + "\">#</span>";

    /// <summary>
    /// The result of a field Word saved without one: text or HTML, or <see cref="WordFieldOutput.None"/> when it cannot be
    /// computed here (its source is not in the document). <paramref name="countedValue"/> is the counter value
    /// <see cref="Count"/> returned for this field.
    /// </summary>
    public static WordFieldOutput Compute(WordFieldInstruction instruction, WordFieldContext context, int? countedValue = null, WordFormFieldState? form = null)
    {
        switch (DisplayOf(instruction.Name))
        {
            case WordFieldDisplay.Hidden: return WordFieldOutput.None;
            case WordFieldDisplay.PagePlaceholder: return new(null, PlaceholderHtml(instruction), false);
        }
        switch (instruction.Name)
        {
            case "EQ": return new(null, WordFieldEquation.ToHtml(EquationCode(instruction.Code)), false);
            case "SYMBOL": return Symbol(instruction);
            case "FORMCHECKBOX":
                return new(null, "<span class=\"form-checkbox\" role=\"img\" aria-label=\"" + (form?.IsChecked == true ? "checked" : "unchecked") + "\">"
                                 + (form?.IsChecked == true ? "☒" : "☐") + "</span>", false);
            case "FORMDROPDOWN":
            {
                string? selected = form?.Entries is { Count: > 0 } entries ? entries[Math.Max(0, Math.Min(form.SelectedEntry, entries.Count - 1))] : null;
                return selected is null ? WordFieldOutput.None : new(selected, null, true);
            }
            case "FORMTEXT": return form?.DefaultText is { } defaultText ? Text(defaultText) : WordFieldOutput.None;
            case "MACROBUTTON": return Text(ButtonText(instruction.Code, wordsBeforeText: 2, wholeRest: true));
            // Word 16: GOTOBUTTON bm Go there → "Go" (only the first word unless quoted)
            case "GOTOBUTTON": return Text(instruction.Arguments.Count > 1 ? instruction.Arguments[1] : "");
            case "ADVANCE": return new(null, AdvanceHtml(instruction), false);
        }
        if (IsCounted(instruction.Name))
        {
            if (countedValue is not { } value) return WordFieldOutput.None;
            if (instruction.Name == "SEQ" && instruction.Has("\\h")) return WordFieldOutput.None;    // hidden, still counted
            return instruction.Name switch
            {
                "LISTNUM" => Text(ListNumberLabel(instruction, value)),
                "AUTONUM" or "AUTONUMLGL" or "AUTONUMOUT" => Text(value.ToString(CultureInfo.InvariantCulture) + (instruction.SwitchArgument("\\s") is { Length: > 0 } separator ? separator : ".")),
                _ => Text(WordFieldFormat.Apply(instruction, new WordFieldValue(Number: value), context)),
            };
        }
        WordFieldValue? computed = Value(instruction, context);
        return computed is { } result ? Text(WordFieldFormat.Apply(instruction, result, context)) : WordFieldOutput.None;
    }

    private static WordFieldOutput Text(string text) => text.Length == 0 ? WordFieldOutput.None : new(text, null, true);

    /// <summary>A field's value before its switches; null when it cannot be computed from the document.</summary>
    private static WordFieldValue? Value(WordFieldInstruction instruction, WordFieldContext context)
    {
        string name = instruction.Name;
        string? argument = instruction.Arguments.Count > 0 ? instruction.Arguments[0] : null;
        switch (name)
        {
            case "=": return WordFieldExpression.Evaluate(instruction.FormulaExpression, context);
            case "IF": return Condition(instruction.Tokens, context, out string? chosen) ? WordFieldValue.OfText(chosen ?? "") : null;
            case "COMPARE": return CompareField(instruction.Tokens, context) is { } compared ? WordFieldValue.OfNumber(compared ? 1 : 0) : null;
            case "QUOTE":
            {
                // Word 16: QUOTE "a" "b" → ab; QUOTE 65 → A (an unquoted number is a character code)
                var plain = instruction.Tokens.Where(t => !t.IsSwitch).ToList();
                if (plain.Count == 0) return null;
                if (plain.Count == 1 && !plain[0].IsQuoted && int.TryParse(plain[0].Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int code) && code is > 0 and < 0x10000)
                    return WordFieldValue.OfText(((char)code).ToString());
                return WordFieldValue.OfText(string.Concat(PlainArguments(instruction)));
            }
            case "MERGEFIELD": return argument is null ? null : WordFieldValue.OfText("«" + argument + "»");     // Word 16: MERGEFIELD Name → «Name»
            case "MERGESEQ": case "MERGEREC": return null;
            case "DATE": return new WordFieldValue(Date: context.Now ?? DateTime.Now, DefaultDatePicture: WordPicture(context.Culture.DateTimeFormat.ShortDatePattern));
            case "TIME": return new WordFieldValue(Date: context.Now ?? DateTime.Now, DefaultDatePicture: WordPicture(context.Culture.DateTimeFormat.ShortTimePattern));
            case "CREATEDATE": case "SAVEDATE": case "PRINTDATE":
                // Word 16 default: the short date and the long time (CREATEDATE → 04/11/2026 10:21:00 on a Hebrew system)
                return context.DocumentDate?.Invoke(name) is { } date
                    ? new WordFieldValue(Date: date, DefaultDatePicture: WordPicture(context.Culture.DateTimeFormat.ShortDatePattern + " " + context.Culture.DateTimeFormat.LongTimePattern))
                    : null;
            case "TITLE": case "SUBJECT": case "AUTHOR": case "KEYWORDS": case "COMMENTS": case "LASTSAVEDBY": case "TEMPLATE":
            case "USERNAME": case "USERINITIALS": case "USERADDRESS":
                // a field with its own text sets the property and shows it (Word 16: TITLE "x" → x)
                if (argument is not null) return WordFieldValue.OfText(argument);
                return context.DocumentProperty?.Invoke(name) is { } property ? WordFieldValue.OfText(property) : null;
            case "REVNUM": case "NUMWORDS": case "NUMCHARS": case "EDITTIME":
                return context.DocumentProperty?.Invoke(name) is { } count && WordFieldFormat.TryParseNumber(count, out decimal number) ? WordFieldValue.OfNumber(number) : null;
            case "INFO":
            {
                if (argument is null) return null;
                var inner = WordFieldInstruction.Parse(string.Join(" ", instruction.Arguments.Select(Quote)));
                return Value(inner, context);
            }
            case "DOCPROPERTY":
                return argument is not null && (context.DocumentProperty?.Invoke(argument.ToUpperInvariant()) ?? context.DocumentProperty?.Invoke(argument)) is { } custom
                    ? WordFieldValue.OfText(custom) : null;
            case "DOCVARIABLE": return argument is not null && context.DocumentVariable?.Invoke(argument) is { } variable ? WordFieldValue.OfText(variable) : null;
            case "FILENAME":
                if (instruction.Has("\\p")) return context.FilePath is { } path ? WordFieldValue.OfText(path) : null;
                return context.FileName is { } fileName ? WordFieldValue.OfText(fileName) : null;
            case "FILESIZE":
            {
                if (context.FileSize is not { } size) return null;
                decimal value = instruction.Has("\\m") ? Math.Round(size / 1048576m, 0, MidpointRounding.AwayFromZero)
                              : instruction.Has("\\k") ? Math.Round(size / 1024m, 0, MidpointRounding.AwayFromZero) : size;
                return WordFieldValue.OfNumber(value);
            }
            case "REF":
                if (argument is null) return null;
                if (instruction.Has("\\n") || instruction.Has("\\r") || instruction.Has("\\w") || instruction.Has("\\p")) return null;   // paragraph number / position: from the result only
                return context.BookmarkText is null ? null : WordFieldValue.OfText(context.BookmarkText(argument) ?? ReferenceNotFound);
            default:
                // { bookmark } is REF bookmark (17.16.5.47: "the REF field instruction may be omitted")
                if (instruction.Arguments.Count == 0 && context.BookmarkText?.Invoke(instruction.Code.Trim().Split(' ')[0]) is { } bookmarkText)
                    return WordFieldValue.OfText(bookmarkText);
                return null;
        }
    }

    private static string Quote(string word) => word.IndexOf(' ') >= 0 ? "\"" + word + "\"" : word;

    private static IEnumerable<string> PlainArguments(WordFieldInstruction instruction)
    {
        foreach (WordFieldToken token in instruction.Tokens) if (!token.IsSwitch) yield return token.Text; else yield break;
    }

    /// <summary>A .NET date pattern as a Word date picture (the same letters, tt → AM/PM).</summary>
    private static string WordPicture(string pattern) => pattern.Replace("tt", "AM/PM").Replace("t", "AM/PM");

    /// <summary>
    /// IF expression1 operator expression2 TrueText FalseText (17.16.5.29). Numbers compare as numbers, anything else as
    /// text (case sensitive, Word 16: IF "abc" = "ABC" → no); with = / &lt;&gt; the right text may hold ? and * wildcards
    /// (Word 16: COMPARE "abc" = "a*" → 1, "a?c" → 1).
    /// </summary>
    private static bool Condition(IReadOnlyList<WordFieldToken> tokens, WordFieldContext context, out string? chosen)
    {
        chosen = null;
        var plain = tokens.TakeWhile(t => !t.IsSwitch).ToList();
        if (plain.Count < 3) return false;
        bool? result = Compare(plain[0], plain[1].Text, plain[2], context);
        if (result is null) return false;
        chosen = result.Value ? (plain.Count > 3 ? plain[3].Text : "") : (plain.Count > 4 ? plain[4].Text : "");
        return true;
    }

    private static bool? CompareField(IReadOnlyList<WordFieldToken> tokens, WordFieldContext context)
    {
        var plain = tokens.TakeWhile(t => !t.IsSwitch).ToList();
        return plain.Count < 3 ? null : Compare(plain[0], plain[1].Text, plain[2], context);
    }

    private static bool? Compare(WordFieldToken left, string op, WordFieldToken right, WordFieldContext context)
    {
        decimal? leftNumber = Operand(left, context), rightNumber = Operand(right, context);
        if (leftNumber is { } a && rightNumber is { } b)
            return op switch { "=" => a == b, "<>" => a != b, "<" => a < b, "<=" => a <= b, ">" => a > b, ">=" => a >= b, _ => null };
        string l = left.Text, r = right.Text;
        return op switch
        {
            "=" => WildcardMatch(l, r),
            "<>" => !WildcardMatch(l, r),
            "<" => string.CompareOrdinal(l, r) < 0, "<=" => string.CompareOrdinal(l, r) <= 0,
            ">" => string.CompareOrdinal(l, r) > 0, ">=" => string.CompareOrdinal(l, r) >= 0,
            _ => null,
        };
    }

    private static decimal? Operand(WordFieldToken token, WordFieldContext context)
    {
        if (WordFieldFormat.TryParseNumber(token.Text, out decimal number)) return number;
        if (token.IsQuoted) return null;
        // an unquoted word may be a bookmark or a formula (IF bm > 3 ...)
        WordFieldValue value = WordFieldExpression.Evaluate(token.Text, context);
        return value.Number;
    }

    private static bool WildcardMatch(string text, string pattern)
    {
        int t = 0, p = 0, starP = -1, starT = 0;
        while (t < text.Length)
        {
            if (p < pattern.Length && (pattern[p] == '?' || pattern[p] == text[t])) { t++; p++; }
            else if (p < pattern.Length && pattern[p] == '*') { starP = p++; starT = t; }
            else if (starP >= 0) { p = starP + 1; t = ++starT; }
            else return false;
        }
        while (p < pattern.Length && pattern[p] == '*') p++;
        return p == pattern.Length;
    }

    /// <summary>MACROBUTTON MacroName DisplayText: the text after the field and macro names, as written (Word 16: "Click here now ").</summary>
    private static string ButtonText(string code, int wordsBeforeText, bool wholeRest)
    {
        int i = 0;
        for (int word = 0; word <= wordsBeforeText; word++)
        {
            while (i < code.Length && char.IsWhiteSpace(code[i])) i++;
            if (word == wordsBeforeText) break;
            while (i < code.Length && !char.IsWhiteSpace(code[i])) i++;
        }
        string rest = i < code.Length ? code.Substring(i) : "";
        if (rest.StartsWith("\"", StringComparison.Ordinal) && rest.IndexOf('"', 1) is int close and > 0) return rest.Substring(1, close - 1);
        return wholeRest ? rest.TrimStart() : rest;
    }

    /// <summary>
    /// LISTNUM's label: NumberDefault 1) a) i) (1) (a) (i) 1. a. i., LegalDefault 1. 1.1. ..., OutlineDefault I. A. 1. a) (1) (a) (i) (a) (i)
    /// - level 1 and 2 of NumberDefault and level 1 of the other two as Word 16 showed them (1), a), 6., VII.); the deeper levels
    /// follow Word's built-in list templates as listed in Microsoft's LISTNUM field help (not re-checked in Word).
    /// </summary>
    private static string ListNumberLabel(WordFieldInstruction instruction, int value)
    {
        string list = instruction.Arguments.Count > 0 ? instruction.Arguments[0] : "NumberDefault";
        int level = instruction.SwitchArgument("\\l") is { Length: > 0 } l && int.TryParse(l, out int parsed) ? Math.Max(1, Math.Min(9, parsed)) : 1;
        if (list.Equals("LegalDefault", StringComparison.OrdinalIgnoreCase)) return value.ToString(CultureInfo.InvariantCulture) + ".";
        if (list.Equals("OutlineDefault", StringComparison.OrdinalIgnoreCase))
            return level switch
            {
                1 => WordFieldFormat.NumberFormat(value, "ROMAN") + ".", 2 => WordFieldFormat.NumberFormat(value, "ALPHABETIC") + ".",
                3 => value + ".", 4 => WordFieldFormat.NumberFormat(value, "alphabetic") + ")", 5 => "(" + value + ")",
                6 => "(" + WordFieldFormat.NumberFormat(value, "alphabetic") + ")", 7 => "(" + WordFieldFormat.NumberFormat(value, "roman") + ")",
                8 => "(" + WordFieldFormat.NumberFormat(value, "alphabetic") + ")", _ => "(" + WordFieldFormat.NumberFormat(value, "roman") + ")",
            };
        return level switch
        {
            1 => value + ")", 2 => WordFieldFormat.NumberFormat(value, "alphabetic") + ")", 3 => WordFieldFormat.NumberFormat(value, "roman") + ")",
            4 => "(" + value + ")", 5 => "(" + WordFieldFormat.NumberFormat(value, "alphabetic") + ")", 6 => "(" + WordFieldFormat.NumberFormat(value, "roman") + ")",
            7 => value + ".", 8 => WordFieldFormat.NumberFormat(value, "alphabetic") + ".", _ => WordFieldFormat.NumberFormat(value, "roman") + ".",
        };
    }

    /// <summary>SYMBOL code \f font \s size \u (17.16.5.60): the character, decimal or 0x hex, in its font.</summary>
    private static WordFieldOutput Symbol(WordFieldInstruction instruction)
    {
        if (instruction.Arguments.Count == 0) return WordFieldOutput.None;
        string number = instruction.Arguments[0];
        bool parsed = number.StartsWith("0x", StringComparison.OrdinalIgnoreCase)
            ? int.TryParse(number.Substring(2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int code)
            : int.TryParse(number, NumberStyles.Integer, CultureInfo.InvariantCulture, out code);
        if (!parsed || code <= 0 || code > 0x10FFFF) return WordFieldOutput.None;
        string font = instruction.SwitchArgument("\\f") ?? "";
        // a symbol font's 8-bit code shows through the private-use mapping Word uses for w:sym (U+F000 + code)
        if (code < 0x100 && font.Length > 0 && IsSymbolFont(font)) code += 0xF000;
        string character = char.ConvertFromUtf32(code);
        var html = new StringBuilder("<span");
        var css = new StringBuilder();
        if (font.Length > 0) css.Append("font-family:'").Append(WordHtmlText.CssFontName(font)).Append('\'');
        if (instruction.SwitchArgument("\\s") is { Length: > 0 } size && decimal.TryParse(size, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal points))
            css.Append(css.Length > 0 ? ";" : "").Append("font-size:").Append(points.ToString(CultureInfo.InvariantCulture)).Append("pt");
        if (css.Length > 0) html.Append(" style=\"").Append(css).Append('"');
        html.Append('>').Append(WordHtmlText.Encode(character)).Append("</span>");
        return new(code >= 0xF000 && code <= 0xF0FF ? null : character, html.ToString(), true);
    }

    private static bool IsSymbolFont(string font) => font.Equals("Symbol", StringComparison.OrdinalIgnoreCase) || font.StartsWith("Wingdings", StringComparison.OrdinalIgnoreCase)
                                                     || font.StartsWith("Webdings", StringComparison.OrdinalIgnoreCase);

    /// <summary>ADVANCE \r n / \l n (points right / left): a gap before the following text; \u / \d / \x / \y move it in a page layout only.</summary>
    private static string AdvanceHtml(WordFieldInstruction instruction)
    {
        decimal Points(string s) => instruction.SwitchArgument(s) is { Length: > 0 } v && decimal.TryParse(v, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal p) ? p : 0;
        decimal offset = Points("\\r") - Points("\\l");
        if (offset == 0) return "";
        return "<span class=\"field-advance\" style=\"display:inline-block;width:0;margin-inline-start:" + offset.ToString(CultureInfo.InvariantCulture) + "pt\"></span>";
    }

    /// <summary>EQ's code after the field name.</summary>
    private static string EquationCode(string code)
    {
        int start = code.IndexOf("EQ", StringComparison.OrdinalIgnoreCase);
        return start < 0 ? "" : code.Substring(start + 2);
    }
}

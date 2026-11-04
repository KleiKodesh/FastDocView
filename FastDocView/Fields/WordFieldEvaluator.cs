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
public readonly record struct WordFieldOutput(string? Text, string? Html, bool IsSearchable, double BaselineShiftPt = 0)
{
    public static readonly WordFieldOutput None = new(null, null, false);
    public bool IsEmpty => Text is null && Html is null && BaselineShiftPt == 0;
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
        // no field can fail the window it is in: a field whose computation throws keeps the result Word saved (None)
        try { return ComputeUnguarded(instruction, context, countedValue, form); }
        catch (Exception error) when (error is not OutOfMemoryException and not WordDocumentSourceChangedException and not OperationCanceledException)
        {
            return WordFieldOutput.None;
        }
    }

    private static WordFieldOutput ComputeUnguarded(WordFieldInstruction instruction, WordFieldContext context, int? countedValue, WordFormFieldState? form)
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
            // \u n / \d n raise / lower the text after the field (17.16.5.2): the readers shift the paragraph's later runs
            case "ADVANCE": return new(null, AdvanceHtml(instruction), false, AdvanceShift(instruction));
            case "NOTEREF": return NoteReference(instruction, context);
            case "TOC": return TableOfContents(instruction, context);
            case "INDEX": return BuiltIndex(instruction, context);
            case "BARCODE": return Barcode(instruction, context);
            case "TOA": return TableOfAuthorities(instruction, context);
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
                return Reference(instruction, argument, context);
            case "STYLEREF": return StyleReference(instruction, context);
            default:
                // { bookmark } is REF bookmark (17.16.5.47: "the REF field instruction may be omitted")
                if (instruction.Arguments.Count == 0 && context.BookmarkText?.Invoke(instruction.Code.Trim().Split(' ')[0]) is { } bookmarkText)
                    return WordFieldValue.OfText(bookmarkText);
                return null;
        }
    }

    /// <summary>Word 16's STYLEREF text when no paragraph has the style (Hebrew UI, as written on the dev machine).</summary>
    public static string StyleNotFound(string style) => "שגיאה! השתמש בכרטיסיה בית כדי להחיל " + style + " על הטקסט שברצונך שיופיע כאן.";

    /// <summary>
    /// REF (17.16.5.47): the bookmark's text; \n \r \w its paragraph's number (Word 16: list label "1." → "1" for all three -
    /// the trailing period dropped; our labels are the shown label, so a multi-level relative / full context is the label
    /// as numbered); \p "above" / "below" (Word 16: \p alone → "above"; with \n → "1 above"; the word follows the text's
    /// language - Hebrew text got "מתחת"; we write English).
    /// </summary>
    private static WordFieldValue? Reference(WordFieldInstruction instruction, string bookmark, WordFieldContext context)
    {
        bool number = instruction.Has("\\n") || instruction.Has("\\r") || instruction.Has("\\w");
        bool position = instruction.Has("\\p");
        bool isListComplete = context.IsBookmarkListComplete?.Invoke() ?? true;
        if (!number && !position)
        {
            if (context.BookmarkText is null) return null;
            return context.BookmarkText(bookmark) is { } text ? WordFieldValue.OfText(text) : isListComplete ? WordFieldValue.OfText(ReferenceNotFound) : null;
        }
        if (context.BookmarkBlock?.Invoke(bookmark) is not { } block) return context.BookmarkBlock is null || !isListComplete ? null : WordFieldValue.OfText(ReferenceNotFound);
        var parts = new List<string>();
        if (number && ParagraphNumber(block, NumberContext(instruction), context) is { Length: > 0 } label) parts.Add(label);
        if (position && Position(block, context) is { } where) parts.Add(where);
        return parts.Count == 0 ? null : WordFieldValue.OfText(string.Join(" ", parts));
    }

    private enum NumberForm : byte { Own, Relative, Full }

    private static NumberForm NumberContext(WordFieldInstruction instruction)
        => instruction.Has("\\w") ? NumberForm.Full : instruction.Has("\\r") ? NumberForm.Relative : NumberForm.Own;

    /// <summary>
    /// A paragraph's number (17.16.5.47 \n \r \w) as Word 16 writes it (COM, 2026-11-04, a three-level list "%1)" "%2)" "%3)" at
    /// 2) b) i)): \n its own level ("i)"); \w full context ("2)b)i)"); \r relative to the field's own paragraph - the levels
    /// they share left out ("i)" from inside 2)b)), and from outside the list its own level ("i)"); a trailing period is
    /// dropped in all three ("1." → "1").
    /// </summary>
    private static string? ParagraphNumber(int block, NumberForm form, WordFieldContext context)
    {
        string? own = context.ListLabelOfBlock?.Invoke(block)?.Trim();
        if (own is null) return null;
        string[]? target = form == NumberForm.Own ? null : context.ListContextOfBlock?.Invoke(block)?.Split('\u0001');
        if (target is not { Length: > 1 }) return own.TrimEnd('.');
        static string Label(string entry) => entry.Substring(entry.IndexOf('\u0002') + 1);
        if (form == NumberForm.Full) return string.Concat(target.Skip(1).Select(Label)).TrimEnd('.');
        string[]? here = context.CurrentBlock is { } current ? context.ListContextOfBlock?.Invoke(current)?.Split('\u0001') : null;
        if (here is not { Length: > 1 } || here[0] != target[0]) return own.TrimEnd('.');
        // the levels the field's own paragraph shares (same level, same number) are left out; the rest from the first that differs
        var hereLevels = new HashSet<string>(here.Skip(1));
        int first = 1;
        while (first < target.Length - 1 && hereLevels.Contains(target[first])) first++;
        return string.Concat(target.Skip(first).Select(Label)).TrimEnd('.');
    }

    /// <summary>\p: "above" / "below" - in Hebrew text "מעל" / "מתחת" (Word 16 wrote those for a Hebrew-language field).</summary>
    private static string? Position(int block, WordFieldContext context)
        => context.CurrentBlock is { } here
            ? (block <= here ? (context.IsRightToLeftText ? "מעל" : "above") : (context.IsRightToLeftText ? "מתחת" : "below"))
            : null;

    /// <summary>NOTEREF (17.16.5.37): the note's mark; \f as a note reference (superscript); \p above / below; \h links (HyperlinkFieldTag).</summary>
    private static WordFieldOutput NoteReference(WordFieldInstruction instruction, WordFieldContext context)
    {
        if (instruction.Arguments.Count == 0 || context.NoteMarkOfBookmark?.Invoke(instruction.Arguments[0]) is not { } mark)
            return context.NoteMarkOfBookmark is null || instruction.Arguments.Count == 0 || !(context.IsBookmarkListComplete?.Invoke() ?? true)
                ? WordFieldOutput.None : new(ReferenceNotFound, null, true);
        string text = mark;
        if (instruction.Has("\\p") && context.BookmarkBlock?.Invoke(instruction.Arguments[0]) is { } block && Position(block, context) is { } where) text += " " + where;
        return instruction.Has("\\f")
            ? new(text, "<span class=\"note-reference\">" + WordHtmlText.Encode(mark) + "</span>" + WordHtmlText.Encode(text.Substring(mark.Length)), true)
            : new(text, null, true);
    }

    /// <summary>
    /// STYLEREF "style" (17.16.5.59): the text of the nearest paragraph in the style - in a flow view the search runs back from
    /// the field's paragraph (Word starts at the page; \l, which starts at the page's bottom, has no page here), then forward;
    /// \n \r \w its number, \p above / below. A heading level number names "heading N" (Word accepts STYLEREF 1).
    /// </summary>
    private static WordFieldValue? StyleReference(WordFieldInstruction instruction, WordFieldContext context)
    {
        if (instruction.Arguments.Count == 0 || context.StyledParagraph is null || context.CurrentBlock is not { } here) return null;
        string style = instruction.Arguments[0];
        if (int.TryParse(style, NumberStyles.Integer, CultureInfo.InvariantCulture, out int level) && level is >= 1 and <= 9) style = "heading " + level;
        (int block, string text)? found = context.StyledParagraph(style, here, false) ?? context.StyledParagraph(style, here, true);
        if (found is not { } paragraph) return WordFieldValue.OfText(StyleNotFound(instruction.Arguments[0]));
        bool number = instruction.Has("\\n") || instruction.Has("\\r") || instruction.Has("\\w");
        if (!number && !instruction.Has("\\p")) return WordFieldValue.OfText(paragraph.text);
        var parts = new List<string>();
        if (number && ParagraphNumber(paragraph.block, NumberContext(instruction), context) is { Length: > 0 } label) parts.Add(label);
        if (instruction.Has("\\p") && Position(paragraph.block, context) is { } where) parts.Add(where);
        return WordFieldValue.OfText(string.Join(" ", parts));
    }

    /// <summary>
    /// A TOC Word saved without a result, built from the document's headings (17.16.5.68): \o "a-b" heading levels (default
    /// 1-9 as Word's \o without levels), \n "a-b" levels without page numbers, page numbers as placeholders (no pages in a
    /// flow view); one line per heading inside the field's block; \h links each entry to its heading's block ("#block-N").
    /// </summary>
    private static WordFieldOutput TableOfContents(WordFieldInstruction instruction, WordFieldContext context)
    {
        if (context.Headings?.Invoke() is not { Count: > 0 } headings) return WordFieldOutput.None;
        (int from, int to) Levels(string? range, (int, int) fallback)
        {
            if (range is not { Length: > 0 }) return fallback;
            string[] ends = range.Split('-');
            return int.TryParse(ends[0], out int a) && int.TryParse(ends[ends.Length - 1], out int b) ? (a, b) : fallback;
        }
        (int from, int to) levels = Levels(instruction.SwitchArgument("\\o"), (1, 9));
        (int from, int to)? noNumbers = instruction.Has("\\n") ? Levels(instruction.SwitchArgument("\\n"), (1, 9)) : null;
        bool links = instruction.Has("\\h");
        var html = new StringBuilder("<span class=\"field-toc\" style=\"display:block\">");
        var text = new StringBuilder();
        foreach (WordDocumentHeading heading in headings)
        {
            if (heading.Level < levels.from || heading.Level > levels.to) continue;
            html.Append("<span class=\"toc-line\" style=\"display:block;padding-inline-start:").Append((heading.Level - 1) * 1.5).Append("em\">");
            // \h: the entry links to its heading's block (the app's FastDocViewPage jumps by "#block-<index>")
            if (links) html.Append("<a class=\"link\" href=\"#block-").Append(heading.BlockIndex.ToString(CultureInfo.InvariantCulture)).Append("\">");
            html.Append(WordHtmlText.Encode(heading.Text));
            if (links) html.Append("</a>");
            if (noNumbers is not { } skip || heading.Level < skip.from || heading.Level > skip.to)
                html.Append("<span class=\"toc-leader\"> </span><span class=\"field-page-placeholder\" title=\"PAGEREF\">#</span>");
            html.Append("</span>");
            text.Append(heading.Text).Append('\n');
        }
        html.Append("</span>");
        return text.Length == 0 ? WordFieldOutput.None : new(text.ToString(), html.ToString(), true);
    }

    /// <summary>
    /// An INDEX Word saved without a result, built from the XE entries (17.16.5.31 INDEX, 17.16.5.76 XE): "main:sub" levels,
    /// sorted (ordinal, case-insensitive); each entry's page numbers as placeholders after \e's separator (", " by default);
    /// XE \t cross-reference text instead of numbers; \f type filters entries (INDEX \f), \b ranges are not separate here.
    /// One line per entry inside the field's block.
    /// </summary>
    private static WordFieldOutput BuiltIndex(WordFieldInstruction instruction, WordFieldContext context)
    {
        if (context.IndexEntries?.Invoke() is not { Count: > 0 } entries) return WordFieldOutput.None;
        string separator = instruction.SwitchArgument("\\e") is { Length: > 0 } e ? e : ", ";
        string? type = instruction.SwitchArgument("\\f");
        var lines = new SortedDictionary<string, (string[] levels, string? crossReference, int count)>(StringComparer.OrdinalIgnoreCase);
        foreach ((_, string code) in entries)
        {
            WordFieldInstruction entry = WordFieldInstruction.Parse(code);
            if (entry.Name != "XE" || entry.Arguments.Count == 0) continue;
            string entryType = entry.SwitchArgument("\\f") is { Length: > 0 } t ? t : "I";      // XE's default type is "I" (17.16.5.76)
            if (!string.Equals(entryType, type ?? "I", StringComparison.OrdinalIgnoreCase)) continue;
            string[] levels = entry.Arguments[0].Split(':').Select(level => level.Trim()).ToArray();
            for (int depth = 1; depth <= levels.Length; depth++)
            {
                string key = string.Join("\u0001", levels.Take(depth));
                lines.TryGetValue(key, out var line);
                lines[key] = (levels.Take(depth).ToArray(), depth == levels.Length ? entry.SwitchArgument("\\t") ?? line.crossReference : line.crossReference,
                              line.count + (depth == levels.Length ? 1 : 0));
            }
        }
        if (lines.Count == 0) return WordFieldOutput.None;
        var html = new StringBuilder("<span class=\"field-index\" style=\"display:block\">");
        var text = new StringBuilder();
        foreach ((string[] levels, string? crossReference, int count) in lines.Values)
        {
            html.Append("<span class=\"index-line\" style=\"display:block;padding-inline-start:").Append((levels.Length - 1) * 1.5).Append("em\">")
                .Append(WordHtmlText.Encode(levels[levels.Length - 1]));
            if (crossReference is not null) html.Append(WordHtmlText.Encode(separator + crossReference));
            else for (int i = 0; i < count; i++) html.Append(WordHtmlText.Encode(i == 0 ? separator : ", ")).Append("<span class=\"field-page-placeholder\" title=\"PAGE\">#</span>");
            html.Append("</span>");
            text.Append(levels[levels.Length - 1]).Append(crossReference is null ? "" : separator + crossReference).Append('\n');
        }
        return new(text.ToString(), html.Append("</span>").ToString(), true);
    }

    /// <summary>
    /// A TOA Word saved without a result, built from the TA entries (17.16.5.69 TOA, 17.16.5.67 TA): the long citations (\l)
    /// of the category TOA \c names (all when absent), sorted, each with page placeholders; one line each inside the block.
    /// </summary>
    private static WordFieldOutput TableOfAuthorities(WordFieldInstruction instruction, WordFieldContext context)
    {
        if (context.IndexEntries?.Invoke() is not { Count: > 0 } entries) return WordFieldOutput.None;
        string? category = instruction.SwitchArgument("\\c");
        var citations = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach ((_, string code) in entries)
        {
            WordFieldInstruction entry = WordFieldInstruction.Parse(code);
            if (entry.Name != "TA" || entry.SwitchArgument("\\l") is not { Length: > 0 } citation) continue;
            if (category is { Length: > 0 } && entry.SwitchArgument("\\c") != category) continue;
            citations.Add(citation);
        }
        if (citations.Count == 0) return WordFieldOutput.None;
        var html = new StringBuilder("<span class=\"field-toa\" style=\"display:block\">");
        foreach (string citation in citations)
            html.Append("<span class=\"toa-line\" style=\"display:block\">").Append(WordHtmlText.Encode(citation))
                .Append("<span class=\"toc-leader\"> </span><span class=\"field-page-placeholder\" title=\"PAGE\">#</span></span>");
        return new(string.Join("\n", citations), html.Append("</span>").ToString(), true);
    }

    /// <summary>
    /// BARCODE (17.16.5.4) as POSTNET bars in SVG: the ZIP / ZIP+4 / delivery-point digits (the argument, or with \b the
    /// bookmark's text; the last 5, 9 or 11 digits found), a check digit making the digit sum a multiple of 10, two-of-five
    /// bars per digit (full bars for the places 7 4 2 1 0; 0 = 11000) between two frame bars - the encoding as Wikipedia's
    /// "POSTNET" gives it from USPS Publication 25 (2003); \f draws a Facing Identification Mark instead.
    /// </summary>
    private static WordFieldOutput Barcode(WordFieldInstruction instruction, WordFieldContext context)
    {
        // \f A..E: a Facing Identification Mark, nine places, 1 = a full bar (A 110010011, B 101101101, C 110101011,
        // D 111010111, E 101000101 - Wikipedia "Facing Identification Mark", from the USPS Domestic Mail Manual 708.9 / 202)
        if (instruction.SwitchArgument("\\f") is { Length: > 0 } fim)
        {
            string? pattern = fim.Trim().ToUpperInvariant() switch
            {
                "A" => "110010011", "B" => "101101101", "C" => "110101011", "D" => "111010111", "E" => "101000101", _ => null,
            };
            if (pattern is null) return WordFieldOutput.None;
            var mark = new StringBuilder("<svg class=\"field-barcode field-fim\" role=\"img\" aria-label=\"FIM ").Append(fim.Trim().ToUpperInvariant())
                .Append("\" viewBox=\"0 0 27 10\" style=\"height:1.2em;width:2.7em;vertical-align:baseline\">");
            for (int i = 0; i < pattern.Length; i++)
                if (pattern[i] == '1') mark.Append("<rect x=\"").Append(i * 3).Append("\" y=\"0\" width=\"1.4\" height=\"10\" fill=\"currentColor\"/>");
            return new(null, mark.Append("</svg>").ToString(), false);
        }
        if (instruction.Arguments.Count == 0) return WordFieldOutput.None;
        string source = instruction.Has("\\b") ? context.BookmarkText?.Invoke(instruction.Arguments[0]) ?? "" : instruction.Arguments[0];
        string digits = new string(source.Where(char.IsDigit).ToArray());
        int length = digits.Length >= 11 ? 11 : digits.Length >= 9 ? 9 : digits.Length >= 5 ? 5 : 0;
        if (length == 0) return WordFieldOutput.None;
        digits = digits.Substring(digits.Length - length);
        int sum = digits.Sum(d => d - '0');
        digits += ((10 - sum % 10) % 10).ToString(CultureInfo.InvariantCulture);
        string[] patterns = ["11000", "00011", "00101", "00110", "01001", "01010", "01100", "10001", "10010", "10100"];
        string bars = "1" + string.Concat(digits.Select(d => patterns[d - '0'])) + "1";
        var svg = new StringBuilder("<svg class=\"field-barcode\" role=\"img\" aria-label=\"").Append(WordHtmlText.Encode(digits))
            .Append("\" viewBox=\"0 0 ").Append(bars.Length * 3).Append(" 10\" style=\"height:0.9em;width:").Append(bars.Length * 0.3).Append("em;vertical-align:baseline\">");
        for (int i = 0; i < bars.Length; i++)
            svg.Append("<rect x=\"").Append(i * 3).Append("\" y=\"").Append(bars[i] == '1' ? 0 : 6).Append("\" width=\"1.4\" height=\"").Append(bars[i] == '1' ? 10 : 4).Append("\" fill=\"currentColor\"/>");
        return new(null, svg.Append("</svg>").ToString(), false);
    }

    private static double AdvanceShift(WordFieldInstruction instruction)
    {
        double Points(string s) => instruction.SwitchArgument(s) is { Length: > 0 } v && double.TryParse(v, NumberStyles.Number, CultureInfo.InvariantCulture, out double p) ? p : 0;
        return Points("\\u") - Points("\\d");
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

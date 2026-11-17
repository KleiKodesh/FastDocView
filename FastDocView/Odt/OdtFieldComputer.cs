using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace FastDocView.Odt;

/// <summary>What <see cref="OdtFieldComputer"/> reads outside the block being rendered - each lookup reads one block (windowed).</summary>
internal interface IOdtFieldLookups
{
    bool IsIndexComplete { get; }
    /// <summary>The block of a bookmark / reference mark, or of a sequence or note target (OdtBodyBlockScanner prefixes); null when not known (yet).</summary>
    int? KnownTargetBlock(string name);
    XElement? ReadBlock(int blockIndex);
    OdtFieldState FieldStateOf(int blockIndex);
    string? ListLabelOf(int blockIndex);
    IReadOnlyList<WordDocumentHeading> KnownHeadings { get; }
    IReadOnlyList<OdtNoteReference> NotesIn(int blockIndex);
    OdtDocumentMeta Meta { get; }
    string FilePath { get; }
    OdtStyleCascade Styles { get; }
}

/// <summary>
/// The fields of an OpenDocument text as LibreOffice prints them (measured 2026-11-17: a .fodt with wrong saved texts exported to
/// PDF by soffice): it recomputes sequences, variables (set / get / user fields / expressions), conditional and hidden text,
/// references (bookmark, reference mark, sequence, note), chapter fields, dates / times that are not fixed and the document
/// information fields; it keeps the saved text of the author / sender / input / placeholder fields and the statistics, and the
/// page fields depend on pages (paged view). A value computed from index-time data (OdtFieldState) is replayed per block; a
/// lookup reads the target block. While a first open is still indexing, a target not found yet shows nothing and marks the
/// block provisional (<see cref="UsedIncompleteIndex"/>), the rule .docx fields follow.
/// </summary>
internal sealed class OdtFieldComputer
{
    private static readonly XNamespace Text = OdtStyleCascade.Text, Office = OdtStyleCascade.Office, Style = OdtStyleCascade.Style;
    private static readonly XNamespace DataStyle = OdtStyleCascade.DataStyle;
    /// <summary>LibreOffice's text for a reference whose target is missing (sw/inc/strings.hrc STR_GETREFFLD_REFITEMNOTFOUND).</summary>
    public const string ReferenceNotFound = "Error: Reference source not found";

    private readonly IOdtFieldLookups _lookups;
    private int _block = -1;
    private XElement? _blockElement;
    private readonly Dictionary<XElement, OdtFieldState> _stateBefore = new();
    private readonly Dictionary<XElement, OdtFieldValue?> _setValues = new();
    private DateTime? _now;

    public OdtFieldComputer(IOdtFieldLookups lookups) => _lookups = lookups;

    /// <summary>Set when a field of the block read data the unfinished index does not have yet (the block is drawn provisionally).</summary>
    public bool UsedIncompleteIndex { get; set; }

    /// <summary>The moment non-fixed dates and times show (LibreOffice: the time of the layout); fixed per computer so a window is consistent.</summary>
    public DateTime Now { get => _now ??= DateTime.Now; set => _now = value; }

    /// <summary>Starts a block: replays its sequences and variable assignments from the state at its start, in document order.</summary>
    public void BeginBlock(int blockIndex, XElement block, OdtFieldState stateAtStart)
    {
        _block = blockIndex;
        _blockElement = block;
        _stateBefore.Clear();
        _setValues.Clear();
        Replay(block, stateAtStart, _stateBefore, _setValues);
    }

    public void EndBlock()
    {
        _block = -1;
        _blockElement = null;
        _stateBefore.Clear();
        _setValues.Clear();
    }

    /// <summary>The same steps as the index pass (OdtBodyBlockScanner.FieldOpened), over a parsed block.</summary>
    private static OdtFieldState Replay(XElement block, OdtFieldState state, Dictionary<XElement, OdtFieldState>? before, Dictionary<XElement, OdtFieldValue?>? values,
                                        Func<XElement, bool>? stopAfter = null)
    {
        foreach (XElement element in block.Descendants())
        {
            if (element.Name.Namespace != Text) continue;
            before?.Add(element, state);
            string local = element.Name.LocalName;
            if (local is not ("sequence" or "variable-set") || (string?)element.Attribute(Text + "name") is not { Length: > 0 } name) continue;
            OdtFieldValue? value;
            if (local == "sequence") state = state.Sequence(name, (string?)element.Attribute(Text + "formula"), out value);
            else
            {
                string? valueType = (string?)element.Attribute(Office + "value-type");
                string? stringValue = (string?)element.Attribute(Office + "string-value");
                string? content = valueType == "string" && stringValue is null ? element.Value : null;
                state = state.SetVariable(name, valueType, (string?)element.Attribute(Text + "formula"), (string?)element.Attribute(Office + "value"),
                                          stringValue, (string?)element.Attribute(Office + "boolean-value"), content, out value);
            }
            values?.Add(element, value);
            if (stopAfter?.Invoke(element) == true) break;
        }
        return state;
    }

    /// <summary>The text a field shows, or null to show the text it was saved with.</summary>
    public string? Compute(XElement field)
    {
        if (field.Name.Namespace != Text || _blockElement is null) return null;
        OdtFieldState state = _stateBefore.TryGetValue(field, out OdtFieldState? known) ? known : OdtFieldState.Empty;
        string? Attribute(string name) => (string?)field.Attribute(Text + name);
        switch (field.Name.LocalName)
        {
            case "sequence":
                return _setValues.TryGetValue(field, out OdtFieldValue? sequence) && sequence is { } number
                    ? FormatSequence(number, (string?)field.Attribute(Style + "num-format")) : null;
            case "variable-set":
                if (Attribute("display") == "none") return "";
                return _setValues.TryGetValue(field, out OdtFieldValue? set) && set is { } setValue ? setValue.Display : null;
            case "variable-get":
            case "user-field-get":
            {
                if (Attribute("display") == "formula") return Attribute("name");
                return Attribute("name") is { Length: > 0 } name && state.ValueOf(name) is { } value ? value.Display : null;
            }
            case "expression":
            {
                if (Attribute("formula") is not { Length: > 0 } formula) return null;
                if (Attribute("display") == "formula") return OdtFieldFormula.StripNamespace(formula);
                return OdtFieldFormula.Evaluate(formula, state.ValueOf)?.Display;
            }
            case "conditional-text":
            {
                if (Attribute("condition") is not { Length: > 0 } condition || OdtFieldFormula.Evaluate(condition, state.ValueOf) is not { } truth) return null;
                return (truth.IsText ? truth.Text!.Length > 0 : truth.Number != 0) ? Attribute("string-value-if-true") ?? "" : Attribute("string-value-if-false") ?? "";
            }
            case "hidden-text":
            {
                // ODF 1.3 part 3 7.7.4: the text is hidden when the condition is true
                if (Attribute("condition") is not { Length: > 0 } condition || OdtFieldFormula.Evaluate(condition, state.ValueOf) is not { } truth) return "";
                return (truth.IsText ? truth.Text!.Length > 0 : truth.Number != 0) ? "" : Attribute("string-value") ?? "";
            }
            case "bookmark-ref": return Reference(field, "", Attribute("ref-name"));
            case "reference-ref": return Reference(field, "", Attribute("ref-name"));
            case "sequence-ref": return Reference(field, OdtBodyBlockScanner.SequenceTargetPrefix, Attribute("ref-name"));
            case "note-ref": return Reference(field, OdtBodyBlockScanner.NoteTargetPrefix, Attribute("ref-name"));
            case "chapter": return Chapter(_block, Attribute("display") ?? "number-and-name", int.TryParse(Attribute("outline-level"), out int level) ? level : 1);
            case "date": return DateField(field, isTime: false);
            case "time": return DateField(field, isTime: true);
            case "title": return Information(field, _lookups.Meta.Title);
            case "subject": return Information(field, _lookups.Meta.Subject);
            case "description": return Information(field, _lookups.Meta.Description);
            case "keywords": return Information(field, string.Join(", ", _lookups.Meta.Keywords));
            case "initial-creator": return Information(field, _lookups.Meta.InitialCreator);
            case "creation-date": return IsFixed(field) ? null : FormatDate(_lookups.Meta.CreationDate, DataStyleOf(field), isTime: false, defaultPattern: "MM/dd/yyyy");
            case "creation-time": return IsFixed(field) ? null : FormatDate(_lookups.Meta.CreationDate, DataStyleOf(field), isTime: true, defaultPattern: "HH:mm:ss");
            case "modification-date": return IsFixed(field) ? null : FormatDate(_lookups.Meta.ModificationDate, DataStyleOf(field), isTime: false, defaultPattern: "MM/dd/yyyy");
            case "modification-time": return IsFixed(field) ? null : FormatDate(_lookups.Meta.ModificationDate, DataStyleOf(field), isTime: true, defaultPattern: "HH:mm:ss");
            case "user-defined":
                if (IsFixed(field)) return null;
                return Attribute("name") is { } property && _lookups.Meta.UserDefined.TryGetValue(property, out string? propertyValue) ? propertyValue : null;
            case "file-name": return IsFixed(field) ? null : FileName(Attribute("display") ?? "full");
            default: return null;
        }
    }

    private static bool IsFixed(XElement field) => (string?)field.Attribute(Text + "fixed") == "true";

    /// <summary>A document information field: the saved text when fixed, else the document's value ("" when the document has none).</summary>
    private static string? Information(XElement field, string? value) => IsFixed(field) ? null : value ?? "";

    private static string FormatSequence(OdtFieldValue value, string? numberFormat)
    {
        double number = value.AsNumber;
        if (number == Math.Floor(number) && number >= 1 && number <= int.MaxValue && !string.IsNullOrEmpty(numberFormat) && numberFormat != "1")
            return OdtListStyle.FormatNumber((int)number, numberFormat!);
        return OdtFieldValue.FormatNumber(number);
    }

    // -- references (sw/source/core/fields/reffld.cxx SwGetRefField::UpdateField) -----------------------------------------

    private string? Reference(XElement field, string prefix, string? refName)
    {
        if (string.IsNullOrEmpty(refName)) return null;
        string format = (string?)field.Attribute(Text + "reference-format") ?? (prefix == OdtBodyBlockScanner.SequenceTargetPrefix ? "page" : "text");
        if (format is "page" or "as-page-style") return null;                 // page numbers: paged view
        int? targetBlock = _lookups.KnownTargetBlock(prefix + refName);
        if (targetBlock is null)
        {
            if (!_lookups.IsIndexComplete) { UsedIncompleteIndex = true; return ""; }
            return ReferenceNotFound;
        }
        XElement? block = targetBlock == _block ? _blockElement : _lookups.ReadBlock(targetBlock.Value);
        if (block is null) return null;
        if (prefix == OdtBodyBlockScanner.NoteTargetPrefix) return NoteLabel(block, targetBlock.Value, refName!);

        // the target mark: a sequence with this ref-name, or the bookmark / reference mark (its start, or the point mark)
        XElement? start = null, end = null;
        if (prefix == OdtBodyBlockScanner.SequenceTargetPrefix)
            start = block.Descendants(Text + "sequence").FirstOrDefault(element => (string?)element.Attribute(Text + "ref-name") == refName);
        else
        {
            bool isBookmark = field.Name.LocalName == "bookmark-ref";
            string point = isBookmark ? "bookmark" : "reference-mark";
            start = block.Descendants().FirstOrDefault(element => element.Name.Namespace == Text && (element.Name.LocalName == point || element.Name.LocalName == point + "-start")
                                                                  && (string?)element.Attribute(Text + "name") == refName);
            if (start?.Name.LocalName == point + "-start")
                end = block.Descendants(Text + point + "-end").FirstOrDefault(element => (string?)element.Attribute(Text + "name") == refName && element.IsAfter(start));
        }
        if (start is null) return _lookups.IsIndexComplete ? ReferenceNotFound : "";
        XElement paragraph = start.AncestorsAndSelf().FirstOrDefault(element => element.Name == Text + "p" || element.Name == Text + "h") ?? block;

        switch (format)
        {
            case "direction":
            {
                // "above" / "below" (LocaleDataWrapper getAboveWord / getBelowWord): the target before the field is above
                bool isAbove = targetBlock < _block || (targetBlock == _block && start.IsBefore(field));
                return isAbove ? "above" : "below";
            }
            case "chapter": return ChapterNumber(targetBlock.Value);
            case "number" or "number-no-superior" or "number-all-superior": return _lookups.ListLabelOf(targetBlock.Value) ?? "";
        }

        if (prefix == OdtBodyBlockScanner.SequenceTargetPrefix)
        {
            OdtFieldState state = targetBlock == _block ? StateAtBlockStart() : _lookups.FieldStateOf(targetBlock.Value);
            var values = new Dictionary<XElement, OdtFieldValue?>();
            Replay(block, state, null, values, element => element == start);
            if (!values.TryGetValue(start, out OdtFieldValue? sequenceValue) || sequenceValue is null) return null;
            string number = FormatSequence(sequenceValue.Value, (string?)start.Attribute(Style + "num-format"));
            if (format == "value") return number;
            (string text, int numberStart, int numberEnd) = ParagraphText(paragraph, start, number);
            string category = (string?)start.Attribute(Text + "name") ?? "";
            int categoryStart = category.Length > 0 ? text.IndexOf(category, StringComparison.Ordinal) : -1;
            int categoryEnd = categoryStart >= 0 ? categoryStart + category.Length : -1;
            switch (format)
            {
                case "category-and-value":
                {
                    int from = categoryStart >= 0 ? Math.Min(numberStart, categoryStart) : numberStart;
                    int to = categoryStart >= 0 ? Math.Max(numberEnd, categoryEnd) : numberEnd;
                    return text.Substring(from, to - from);
                }
                case "caption":
                {
                    // SwGetExpField::GetReferenceTextPos: after the category and number, a non-alphanumeric character and the blanks after it are skipped
                    int from = categoryStart >= 0 ? Math.Max(numberEnd, categoryEnd) : numberEnd;
                    if (from < text.Length && (!char.IsLetterOrDigit(text[from]) || text[from] is ' ' or '\t'))
                    {
                        from++;
                        while (from < text.Length && text[from] is ' ' or '\t') from++;
                    }
                    return from < text.Length ? text.Substring(from) : "";
                }
                default: return text;          // "text": the whole paragraph
            }
        }

        // bookmark / reference mark text: from its start to its end, or to the end of the paragraph when the end is elsewhere
        return MarkedText(paragraph, start, end);
    }

    private OdtFieldState StateAtBlockStart()
        => _blockElement is not null && _blockElement.Descendants().FirstOrDefault(element => element.Name.Namespace == Text) is { } first
           && _stateBefore.TryGetValue(first, out OdtFieldState? state) ? state : _lookups.FieldStateOf(_block);

    /// <summary>A note's shown number (GetViewNumStr): the label the index gave the note with that text:id, by its place among the block's notes.</summary>
    private string? NoteLabel(XElement block, int blockIndex, string id)
    {
        int ordinal = 0;
        foreach (XElement note in block.Descendants(Text + "note"))
        {
            if (note.Ancestors(Text + "note-body").Any()) continue;
            if ((string?)note.Attribute(Text + "id") == id)
            {
                IReadOnlyList<OdtNoteReference> notes = _lookups.NotesIn(blockIndex);
                if (ordinal >= notes.Count) return null;
                // the note's number LibreOffice counts (it renumbers notes when it prints: a saved citation is not kept); a custom mark is its label
                OdtNoteReference reference = notes[ordinal];
                return reference.Number == 0 ? reference.Label
                    : OdtListStyle.FormatNumber(reference.Number, _lookups.Styles.NotesConfiguration(reference.Kind).NumberFormat);
            }
            ordinal++;
        }
        return _lookups.IsIndexComplete ? ReferenceNotFound : "";
    }

    /// <summary>The paragraph's text as shown (fields by their saved text, the target sequence by <paramref name="number"/>) and where that number stands.</summary>
    private static (string text, int numberStart, int numberEnd) ParagraphText(XElement paragraph, XElement sequence, string number)
    {
        var text = new StringBuilder();
        int numberStart = -1, numberEnd = -1;
        AppendText(paragraph, text, element =>
        {
            if (element != sequence) return null;
            numberStart = text.Length;
            numberEnd = numberStart + number.Length;
            return number;
        });
        if (numberStart < 0) { numberStart = numberEnd = 0; }
        return (text.ToString(), numberStart, numberEnd);
    }

    private static string MarkedText(XElement paragraph, XElement start, XElement? end)
    {
        if (start.Name.LocalName is "bookmark" or "reference-mark") return "";      // a point mark has no text
        var text = new StringBuilder();
        bool isInside = false, isDone = false;
        void Walk(XElement container)
        {
            foreach (XNode node in container.Nodes())
            {
                if (isDone) return;
                if (node is XText value) { if (isInside) text.Append(value.Value); continue; }
                if (node is not XElement element) continue;
                if (element == start) { isInside = true; continue; }
                if (element == end) { isDone = true; return; }
                if (!isInside) { if (element.Name.Namespace == Text && !IsSkippedInText(element)) Walk(element); continue; }
                AppendElementText(element, text, null);
            }
        }
        Walk(paragraph);
        return text.ToString();
    }

    private static bool IsSkippedInText(XElement element)
        => element.Name.LocalName is "note" or "note-body" or "note-citation" or "ruby-text" or "hidden-text" or "tracked-changes";

    private static void AppendText(XElement container, StringBuilder text, Func<XElement, string?>? replace)
    {
        foreach (XNode node in container.Nodes())
        {
            if (node is XText value) text.Append(value.Value);
            else if (node is XElement element) AppendElementText(element, text, replace);
        }
    }

    private static void AppendElementText(XElement element, StringBuilder text, Func<XElement, string?>? replace)
    {
        if (replace?.Invoke(element) is { } replaced) { text.Append(replaced); return; }
        if (element.Name.Namespace != Text) return;                     // drawings, annotations
        switch (element.Name.LocalName)
        {
            case "s": text.Append(' ', int.TryParse((string?)element.Attribute(Text + "c"), out int count) && count > 0 ? Math.Min(count, 1000) : 1); return;
            case "tab": text.Append('\t'); return;
            case "line-break": text.Append('\n'); return;
        }
        if (IsSkippedInText(element)) return;
        AppendText(element, text, replace);
    }

    // -- chapter (sw/source/core/fields/chpfld.cxx) -------------------------------------------------------------------------

    /// <summary>The nearest heading at or before the block whose outline level is at most <paramref name="level"/> (FindOutlineNodeOfLevel).</summary>
    private int? ChapterHeading(int blockIndex, int level)
    {
        IReadOnlyList<WordDocumentHeading> headings = _lookups.KnownHeadings;
        int low = 0, high = headings.Count;
        while (low < high) { int middle = (low + high) / 2; if (headings[middle].BlockIndex <= blockIndex) low = middle + 1; else high = middle; }
        for (int i = low - 1; i >= 0; i--)
            if (headings[i].Level <= level) return headings[i].BlockIndex;
        return null;
    }

    private string Chapter(int blockIndex, string display, int level)
    {
        if (ChapterHeading(blockIndex, Math.Max(1, level)) is not { } heading) return "";
        XElement? element = heading == _block ? _blockElement : _lookups.ReadBlock(heading);
        if (element is null) return "";
        string number = element.Element(Text + "number")?.Value ?? "";
        var title = new StringBuilder();
        AppendText(element, title, child => child.Name == Text + "number" ? "" : null);
        string name = title.ToString().Trim();
        string plain = number.Trim().TrimEnd('.', ')', ' ').TrimStart('(');
        return display switch
        {
            "name" => name,
            "number" => number,
            "plain-number" => plain,
            "plain-number-and-name" => plain.Length > 0 ? plain + " " + name : name,
            _ => number.Length > 0 ? number + " " + name : name,
        };
    }

    private string ChapterNumber(int blockIndex)
    {
        if (ChapterHeading(blockIndex, 10) is not { } heading) return "";
        XElement? element = heading == _block ? _blockElement : _lookups.ReadBlock(heading);
        return element?.Element(Text + "number")?.Value ?? "";
    }

    // -- dates, times, file name --------------------------------------------------------------------------------------------

    private XElement? DataStyleOf(XElement field) => _lookups.Styles.DataStyleNamed((string?)field.Attribute(Style + "data-style-name"));

    private string? DateField(XElement field, bool isTime)
    {
        DateTime? value;
        if (IsFixed(field))
        {
            value = isTime ? OdtDocumentMeta.ParseTime((string?)field.Attribute(Text + "time-value")) : OdtDocumentMeta.ParseDate((string?)field.Attribute(Text + "date-value"));
            if (value is null) return null;
        }
        else value = Now;
        // text:date-adjust / text:time-adjust (19.873 / 19.880): an xsd:duration added to the value
        if ((string?)field.Attribute(Text + (isTime ? "time-adjust" : "date-adjust")) is { Length: > 0 } adjust)
        {
            try { value = value.Value + System.Xml.XmlConvert.ToTimeSpan(adjust); } catch (FormatException) { } catch (OverflowException) { } catch (ArgumentOutOfRangeException) { }
        }
        return FormatDate(value, DataStyleOf(field), isTime, isTime ? "HH:mm:ss" : "M/d/yy");
    }

    private string FileName(string display)
    {
        string path = _lookups.FilePath;
        return display switch
        {
            "name" => Path.GetFileNameWithoutExtension(path),
            "name-and-extension" => Path.GetFileName(path),
            "path" => Path.GetDirectoryName(path) ?? "",
            _ => path,
        };
    }

    /// <summary>
    /// A date / time in its number:date-style or number:time-style (ODF 1.3 part 3 16.29.11 / 16.29.19: number:day, month, year,
    /// day-of-week, hours, minutes, seconds, am-pm, text); without a data style, LibreOffice's default for the field (measured).
    /// </summary>
    internal static string FormatDate(DateTime? value, XElement? style, bool isTime, string defaultPattern)
    {
        if (value is not { } date) return "";
        CultureInfo culture = CultureOf(style);
        if (style is null) return date.ToString(defaultPattern, CultureInfo.InvariantCulture);
        bool hasAmPm = style.Element(DataStyle + "am-pm") is not null;
        var text = new StringBuilder();
        foreach (XElement part in style.Elements())
        {
            if (part.Name.Namespace != DataStyle) continue;
            bool isLong = (string?)part.Attribute(DataStyle + "style") == "long";
            switch (part.Name.LocalName)
            {
                case "day": text.Append(date.Day.ToString(isLong ? "00" : "0", CultureInfo.InvariantCulture)); break;
                case "month":
                    if ((string?)part.Attribute(DataStyle + "textual") == "true")
                        text.Append(isLong ? culture.DateTimeFormat.GetMonthName(date.Month) : culture.DateTimeFormat.GetAbbreviatedMonthName(date.Month));
                    else text.Append(date.Month.ToString(isLong ? "00" : "0", CultureInfo.InvariantCulture));
                    break;
                case "year": text.Append(isLong ? date.Year.ToString("0000", CultureInfo.InvariantCulture) : (date.Year % 100).ToString("00", CultureInfo.InvariantCulture)); break;
                case "day-of-week": text.Append(isLong ? culture.DateTimeFormat.GetDayName(date.DayOfWeek) : culture.DateTimeFormat.GetAbbreviatedDayName(date.DayOfWeek)); break;
                case "hours":
                {
                    int hour = hasAmPm ? (date.Hour % 12 == 0 ? 12 : date.Hour % 12) : date.Hour;
                    text.Append(hour.ToString(isLong ? "00" : "0", CultureInfo.InvariantCulture));
                    break;
                }
                case "minutes": text.Append(date.Minute.ToString(isLong ? "00" : "0", CultureInfo.InvariantCulture)); break;
                case "seconds": text.Append(date.Second.ToString(isLong ? "00" : "0", CultureInfo.InvariantCulture)); break;
                case "am-pm": text.Append(date.Hour < 12 ? culture.DateTimeFormat.AMDesignator : culture.DateTimeFormat.PMDesignator); break;
                case "text": text.Append(part.Value); break;
            }
        }
        return text.ToString();
    }

    private static CultureInfo CultureOf(XElement? style)
    {
        string? language = (string?)style?.Attribute(DataStyle + "language");
        if (string.IsNullOrEmpty(language)) return CultureInfo.GetCultureInfo("en-US");
        string? country = (string?)style!.Attribute(DataStyle + "country");
        try { return CultureInfo.GetCultureInfo(string.IsNullOrEmpty(country) ? language : language + "-" + country); }
        catch (CultureNotFoundException) { return CultureInfo.GetCultureInfo("en-US"); }
    }
}

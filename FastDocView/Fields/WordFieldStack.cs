using System.Text;

namespace FastDocView.Fields;

/// <summary>Where a piece of text inside fields goes.</summary>
public enum WordFieldTextRoute : byte
{
    /// <summary>Shown (and searched): outside fields, or in a result that shows.</summary>
    Visible,
    /// <summary>Part of a field's instruction (a field code, or a nested field's result inside an instruction): not shown.</summary>
    Instruction,
    /// <summary>A result the reader replaces (page placeholder, hidden field, a recomputed result): not shown.</summary>
    Suppressed,
}

/// <summary>
/// The open fields at the current position, for every reader of every format (.docx w:fldChar begin / separate / end and
/// w:fldSimple, .doc 0x13 / 0x14 / 0x15, .rtf \field / \fldinst / \fldrslt) - one state machine so the renderer and both
/// search readers agree on what shows. Word's model (ECMA-376 Part 1 17.16.2; [MS-DOC] 2.8.25 PlcFld): a field's code runs
/// from begin to separate, its cached result from separate to end; a field nested in another field's code contributes its
/// result to that code ({ IF { PAGE } = 1 ... }); a field ended without a separator has no result.
/// </summary>
public sealed class WordFieldStack
{
    public sealed class Field
    {
        public readonly StringBuilder Instruction = new();
        public bool IsSeparated;
        public bool IsDirty;
        public bool IsLocked;
        public WordFormFieldState? Form;
        /// <summary>The cached result is not shown (placeholder, hidden, recomputed, or inside another field's code).</summary>
        public bool SuppressesResult;
        /// <summary>The result belongs to an enclosing field's instruction.</summary>
        public bool IsInsideInstruction;
        public bool IsCounted;
        public string? LinkTag;
        public WordFieldInstruction? Parsed;
    }

    private readonly List<Field> _fields = new();

    /// <summary>An index entry holding the counters (<see cref="WordFieldSequenceCounters.Serialize"/>) rather than an open field.</summary>
    public const char CountersEntryPrefix = 'C';

    public WordFieldContext Context { get; }

    /// <summary>The counters SEQ / AUTONUM / LISTNUM advance; null: they are not counted here (results stay cached).</summary>
    public WordFieldSequenceCounters? Counters { get; set; }

    /// <summary>Whether this reader computes results (false: a search reader that only needs routing would still compute - kept true by all readers so search sees what shows).</summary>
    public Func<string, string?>? LinkTagOf { get; init; }

    public WordFieldStack(WordFieldContext? context = null) { Context = context ?? new WordFieldContext(); }

    public int Depth => _fields.Count;
    public IReadOnlyList<Field> Fields => _fields;
    public void Clear() => _fields.Clear();

    /// <summary>
    /// Restores the fields open where a block starts, as the index records them (outermost first; 'S' separated / 'I' in its
    /// instruction, then the instruction).
    /// </summary>
    public void Restore(string? openFields)
    {
        _fields.Clear();
        if (Counters is not null) Counters = new WordFieldSequenceCounters();
        if (string.IsNullOrEmpty(openFields)) return;
        foreach (string open in openFields!.Split('\u0001'))
        {
            if (open.Length > 0 && open[0] == CountersEntryPrefix)
            {
                // the counters where the block starts; a reader that does not count leaves them
                if (Counters is not null) Counters = WordFieldSequenceCounters.Deserialize(open.Substring(1));
                continue;
            }
            var field = new Field();
            if (open.Length > 0) field.Instruction.Append(open, 1, open.Length - 1);
            _fields.Add(field);
            if (open.Length > 0 && open[0] == 'S') SeparateRestored(field);
        }
    }

    /// <summary>A restored separated field: its result mode as at its separator, without output (that was in an earlier block) and without counting.</summary>
    private void SeparateRestored(Field field)
    {
        field.IsSeparated = true;
        field.Parsed = WordFieldInstruction.Parse(field.Instruction.ToString());
        field.IsInsideInstruction = IsInsideInstruction(_fields.Count - 1);
        if (field.IsInsideInstruction) return;
        field.SuppressesResult = WordFieldEvaluator.DisplayOf(field.Parsed.Name) != WordFieldDisplay.Result;
        field.LinkTag = LinkTagOf?.Invoke(field.Instruction.ToString());
    }

    public void Begin(bool isDirty = false, bool isLocked = false, WordFormFieldState? form = null)
        => _fields.Add(new Field { IsDirty = isDirty, IsLocked = isLocked, Form = form });

    private bool IsInsideInstruction(int index)
    {
        for (int i = index - 1; i >= 0; i--) if (!_fields[i].IsSeparated) return true;
        return false;
    }

    /// <summary>Where text at the current position goes; instruction text is appended to its field.</summary>
    public WordFieldTextRoute Route(string? text = null)
    {
        for (int i = _fields.Count - 1; i >= 0; i--)
            if (!_fields[i].IsSeparated)
            {
                if (text is not null) _fields[i].Instruction.Append(text);
                return WordFieldTextRoute.Instruction;
            }
        foreach (Field field in _fields) if (field.SuppressesResult) return WordFieldTextRoute.Suppressed;
        return WordFieldTextRoute.Visible;
    }

    /// <summary>Instruction text (w:instrText, an RTF \fldinst, .doc characters between 0x13 and 0x14).</summary>
    public void AppendInstruction(string text)
    {
        if (_fields.Count > 0 && !_fields[_fields.Count - 1].IsSeparated) _fields[_fields.Count - 1].Instruction.Append(text);
    }

    /// <summary>The innermost link whose result shows here (HYPERLINK, REF \h ...).</summary>
    public string? ActiveLink
    {
        get
        {
            for (int i = _fields.Count - 1; i >= 0; i--)
                if (_fields[i].IsSeparated && !_fields[i].IsInsideInstruction && _fields[i].LinkTag is { } tag) return tag;
            return null;
        }
    }

    /// <summary>The separator: the code is complete. Returns what shows in place of the cached result (empty: the cached result shows).</summary>
    public WordFieldOutput Separate()
    {
        if (_fields.Count == 0) return WordFieldOutput.None;
        Field field = _fields[_fields.Count - 1];
        if (field.IsSeparated) return WordFieldOutput.None;
        field.IsSeparated = true;
        return Finish(field, hasResult: true);
    }

    /// <summary>The end of the innermost field. Returns what shows for a field that had no separator (computed / placeholder).</summary>
    public WordFieldOutput End(out string? linkTag)
    {
        linkTag = null;
        if (_fields.Count == 0) return WordFieldOutput.None;
        Field field = _fields[_fields.Count - 1];
        WordFieldOutput output = WordFieldOutput.None;
        if (!field.IsSeparated)
        {
            output = Finish(field, hasResult: false);
            linkTag = field.LinkTag;
        }
        _fields.RemoveAt(_fields.Count - 1);
        // a computed result inside another field's code is part of that code
        if (field.IsInsideInstruction)
        {
            if (output.Text is { } text) Route(text);
            return WordFieldOutput.None;
        }
        return output;
    }

    /// <summary>Whether output of the innermost field (its placeholder / computed result) shows: no field around it is in its code or replaces its result.</summary>
    public bool OutputShows
    {
        get
        {
            for (int i = 0; i < _fields.Count - 1; i++) if (!_fields[i].IsSeparated || _fields[i].SuppressesResult) return false;
            return true;
        }
    }

    /// <summary>The link a computed output (returned by <see cref="Separate"/>) is wrapped in: the field's own.</summary>
    public string? CurrentFieldLink => _fields.Count > 0 ? _fields[_fields.Count - 1].LinkTag : null;

    private WordFieldOutput Finish(Field field, bool hasResult)
    {
        string code = field.Instruction.ToString();
        WordFieldInstruction instruction = field.Parsed = WordFieldInstruction.Parse(code);
        field.IsInsideInstruction = IsInsideInstruction(_fields.IndexOf(field));
        int? counted = null;
        if (Counters is not null && WordFieldEvaluator.IsCounted(instruction.Name)) counted = WordFieldEvaluator.Count(instruction, Counters);
        if (field.IsInsideInstruction)
        {
            // the cached result (or a computed one at End) becomes part of the enclosing code
            if (!hasResult && !field.IsLocked) return WordFieldEvaluator.Compute(instruction, Context, counted, field.Form) is { Text: { } t } ? new(t, null, true) : WordFieldOutput.None;
            return WordFieldOutput.None;
        }
        field.LinkTag = LinkTagOf?.Invoke(code);
        WordFieldDisplay display = WordFieldEvaluator.DisplayOf(instruction.Name);
        if (display != WordFieldDisplay.Result)
        {
            field.SuppressesResult = true;
            return display == WordFieldDisplay.PagePlaceholder ? new(null, WordFieldEvaluator.PlaceholderHtml(instruction), false) : WordFieldOutput.None;
        }
        // Word keeps a locked field's result as it is (17.16.2 fldLock); otherwise a field with no result, one drawn from its
        // code, or one marked dirty (Word recalculates it on open, 17.16.18 dirty) is computed where the document holds its source
        bool compute = !field.IsLocked && (!hasResult || field.IsDirty || WordFieldEvaluator.IsDrawnFromCode(instruction.Name));
        if (!compute) return WordFieldOutput.None;
        WordFieldOutput computed = WordFieldEvaluator.Compute(instruction, Context, counted, field.Form);
        if (computed.IsEmpty)
        {
            // nothing computed: a field drawn from its code shows nothing else (its cached result is empty in Word too)
            if (WordFieldEvaluator.IsDrawnFromCode(instruction.Name) && instruction.Name is not "SYMBOL") field.SuppressesResult = hasResult;
            return WordFieldOutput.None;
        }
        field.SuppressesResult = true;
        return computed;
    }
}

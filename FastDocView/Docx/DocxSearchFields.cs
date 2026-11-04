using FastDocView.Fields;

namespace FastDocView.Docx;

/// <summary>
/// Fields for the .docx search readers (<see cref="DocxSearchTextScanner"/> and the reference
/// <see cref="WordXmlSearchText"/>): the same <see cref="WordFieldStack"/> the renderer runs, so the search reads what shows -
/// a cached result once, a computed result instead of a missing / recomputed one, never a field code, never a page
/// placeholder or a hidden field's text. A text box (w:txbxContent) starts its own fields, as in the renderer; the
/// counters (SEQ ...) are shared with it.
/// </summary>
internal sealed class DocxSearchFields
{
    private readonly WordFieldContext? _context;
    private readonly List<(int depth, WordFieldStack outer)> _textBoxes = new();
    private WordFieldStack _stack;
    private readonly bool _counts;

    /// <param name="counts">Whether SEQ / AUTONUM / LISTNUM count here (body blocks: yes, from the index's counters; notes: no, as the note renderer).</param>
    public DocxSearchFields(WordFieldContext? context, bool counts = true)
    {
        _context = context;
        _counts = counts;
        _stack = New(counts ? new WordFieldSequenceCounters() : null);
    }

    private WordFieldStack New(WordFieldSequenceCounters? counters) => new(_context) { Counters = counters };

    /// <summary>The fields open where a block starts and the counters there (the index's open-field entry), as the renderer restores them.</summary>
    public void Restore(string? openFields)
    {
        _textBoxes.Clear();
        _stack = New(_counts ? new WordFieldSequenceCounters() : null);
        _stack.Restore(openFields);
    }

    /// <summary>True when no field is open (the common case: text is fed as is).</summary>
    public bool IsIdle => _stack.Depth == 0;

    public void Begin(bool isDirty, bool isLocked, WordFormFieldState? form) => _stack.Begin(isDirty, isLocked, form);

    public void Instruction(string text) => _stack.AppendInstruction(text);

    /// <summary>Whether text here shows; instruction text is kept for its field.</summary>
    public bool Shows(string? text) => _stack.Depth == 0 || _stack.Route(text) == WordFieldTextRoute.Visible;

    public void Separate(WordDocumentSearchMatcher matcher)
    {
        WordFieldOutput output = _stack.Separate();
        if (output is { Text: { } text, IsSearchable: true } && _stack.OutputShows) matcher.Feed(text.AsSpan());
    }

    public void End(WordDocumentSearchMatcher matcher)
    {
        WordFieldOutput output = _stack.End(out _);
        if (output is { Text: { } text, IsSearchable: true } && _stack.Route() == WordFieldTextRoute.Visible) matcher.Feed(text.AsSpan());
    }

    public void EnterTextBox(int depth)
    {
        _textBoxes.Add((depth, _stack));
        _stack = New(_stack.Counters);
    }

    /// <summary>An element closed at <paramref name="depth"/>: a text box ending there gives the outer fields back.</summary>
    public void Closed(int depth)
    {
        if (_textBoxes.Count > 0 && _textBoxes[_textBoxes.Count - 1].depth == depth)
        {
            WordFieldSequenceCounters? counters = _stack.Counters;
            _stack = _textBoxes[_textBoxes.Count - 1].outer;
            _stack.Counters = counters;
            _textBoxes.RemoveAt(_textBoxes.Count - 1);
        }
    }

    /// <summary>
    /// A form field's state from the w:ffData elements met inside a w:fldChar (17.16.17), collected element by element by
    /// the byte scanner: the same reading as the renderer's (w:checked over w:default; w:result over w:default).
    /// </summary>
    public sealed class FormDataReader
    {
        private string? _kind;
        private string? _checked, _checkDefault, _listResult, _listDefault, _textDefault;
        private bool _hasChecked, _hasCheckDefault;
        private readonly List<string> _entries = new();
        private string? _parent;

        public void Element(string localName, string? value, string? parentKind)
        {
            switch (localName)
            {
                case "checkBox": _kind = "checkBox"; break;
                case "ddList": _kind = "ddList"; break;
                case "textInput": _kind = "textInput"; break;
                case "listEntry": _entries.Add(value ?? ""); break;
                case "checked": _hasChecked = true; _checked = value; break;
                case "result": _listResult = value; break;
                case "default":
                    if (parentKind == "checkBox") { _hasCheckDefault = true; _checkDefault = value; }
                    else if (parentKind == "ddList") _listDefault = value;
                    else if (parentKind == "textInput") _textDefault = value;
                    break;
            }
            _parent = parentKind;
        }

        public WordFormFieldState? State()
        {
            switch (_kind)
            {
                case "checkBox":
                {
                    bool has = _hasChecked || _hasCheckDefault;
                    string? value = _hasChecked ? _checked : _checkDefault;
                    return new WordFormFieldState { IsChecked = has && value is null or "1" or "true" or "on" };
                }
                case "ddList":
                    return new WordFormFieldState { Entries = _entries.ToList(), SelectedEntry = int.TryParse(_listResult ?? _listDefault, out int s) ? s : 0 };
                case "textInput": return new WordFormFieldState { DefaultText = _textDefault };
                default: return null;
            }
        }
    }
}

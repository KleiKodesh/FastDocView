using System.Globalization;
using System.Text;
using FastDocView.Fields;

namespace FastDocView.Rtf;

/// <summary>
/// RTF fields (RTF 1.5 spec "Fields": {\field &lt;fieldmod&gt;? {\*\fldinst ...} {\fldrslt ...}}) on the shared
/// <see cref="WordFieldStack"/>: \field begins a field, the instruction arrives when its group closes (a nested field there
/// comes in as its result text - the parser captures it into the code), \fldrslt separates, the \field group's close ends
/// it. \fldlock locks it; \fldpriv ("result is not in a form suitable for display") hides its result. \flddirty is not a
/// recalculation request in RTF ("a formatting change has been made to the field result") and is not read. The index
/// pass, its span rescans and the renderer all run one tracker, so a block starts with the fields and counters that were
/// open there (<see cref="Serialize"/> / <see cref="Restore"/>).
/// </summary>
internal sealed class RtfFieldTracker
{
    private readonly WordFieldStack _stack;
    private readonly List<int> _depths = new();          // per open field: its \field group's depth

    public RtfFieldTracker(WordFieldContext? context, bool counts)
    {
        _stack = new WordFieldStack(context) { LinkTagOf = DocxHtml.DocxBlockHtmlRenderer.HyperlinkFieldTag, Counters = counts ? new WordFieldSequenceCounters() : null };
    }

    public int Depth => _stack.Depth;
    public bool IsIdle => _stack.Depth == 0;
    public string? ActiveLink => _stack.ActiveLink;
    public WordFieldStack.Field? Innermost => _stack.Depth > 0 ? _stack.Fields[_stack.Depth - 1] : null;

    public void Start(int depth)
    {
        _stack.Begin();
        _depths.Add(depth);
    }

    public void Lock() { if (Innermost is { IsSeparated: false } field) field.IsLocked = true; }
    public void Private() { if (Innermost is { } field) field.IsPrivateResult = true; }
    public void Form(WordFormFieldState state) { if (Innermost is { IsSeparated: false } field) field.Form = state; }
    public void Instruction(string text) => _stack.AppendInstruction(text);

    /// <summary>\fldrslt: what shows instead of the cached result (empty: the result shows), and whether it shows here.</summary>
    public WordFieldOutput ResultStart(out string? link)
    {
        WordFieldOutput output = _stack.Separate();
        link = _stack.CurrentFieldLink;
        return _stack.OutputShows ? output : WordFieldOutput.None;
    }

    /// <summary>A group closed at <paramref name="depth"/>: the fields whose \field group it was end; their outputs (fields without a result) in order.</summary>
    public void GroupClosed(int depth, Action<WordFieldOutput, string?> output)
    {
        while (_depths.Count > 0 && _depths[_depths.Count - 1] > depth)
        {
            _depths.RemoveAt(_depths.Count - 1);
            WordFieldOutput ended = _stack.End(out string? link);
            if (!ended.IsEmpty && _stack.Route() == WordFieldTextRoute.Visible) output(ended, link);
        }
    }

    /// <summary>Whether text at this point shows (outside fields, or in a result that shows).</summary>
    public bool Shows(ReadOnlySpan<char> text) => _stack.Depth == 0 || _stack.Route(text.ToString()) == WordFieldTextRoute.Visible;

    /// <summary>
    /// The open fields and counters as one string: per field "depth:S|I:L|-:P|-:instruction" joined by U+0001, then
    /// "C" + the counters - what a block start needs (the index stores it per block).
    /// </summary>
    public string Serialize()
    {
        var text = new StringBuilder();
        for (int i = 0; i < _stack.Depth; i++)
        {
            WordFieldStack.Field field = _stack.Fields[i];
            if (text.Length > 0) text.Append('\u0001');
            text.Append(_depths[i].ToString(CultureInfo.InvariantCulture)).Append(':').Append(field.IsSeparated ? 'S' : 'I')
                .Append(field.IsLocked ? 'L' : '-').Append(field.IsPrivateResult ? 'P' : '-').Append(':').Append(field.Instruction);
        }
        if (_stack.Counters is { IsEmpty: false } counters)
        {
            if (text.Length > 0) text.Append('\u0001');
            text.Append(WordFieldStack.CountersEntryPrefix).Append(counters.Serialize());
        }
        return text.ToString();
    }

    /// <summary>Rebuilds the state <see cref="Serialize"/> wrote (no outputs: they were in an earlier block).</summary>
    public void Restore(string? state)
    {
        _stack.Clear();
        _depths.Clear();
        if (_stack.Counters is not null) _stack.Counters = new WordFieldSequenceCounters();
        if (string.IsNullOrEmpty(state)) return;
        foreach (string entry in state!.Split('\u0001'))
        {
            if (entry.Length > 0 && entry[0] == WordFieldStack.CountersEntryPrefix)
            {
                if (_stack.Counters is not null) _stack.Counters = WordFieldSequenceCounters.Deserialize(entry.Substring(1));
                continue;
            }
            int colon = entry.IndexOf(':');
            if (colon <= 0 || colon + 4 > entry.Length || !int.TryParse(entry.Substring(0, colon), NumberStyles.Integer, CultureInfo.InvariantCulture, out int depth)) continue;
            bool separated = entry[colon + 1] == 'S';
            WordFieldSequenceCounters? counters = _stack.Counters;
            _stack.Counters = null;                            // a restored field was counted where it began
            _stack.Begin(isLocked: entry[colon + 2] == 'L');
            _depths.Add(depth);
            if (entry[colon + 3] == 'P') _stack.Fields[_stack.Depth - 1].IsPrivateResult = true;
            _stack.AppendInstruction(entry.Substring(colon + 5));
            if (separated) _stack.Separate();
            _stack.Counters = counters;
        }
    }
}

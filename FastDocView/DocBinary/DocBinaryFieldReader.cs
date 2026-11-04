using FastDocView.Fields;

namespace FastDocView.DocBinary;

/// <summary>
/// .doc fields character by character, for the renderer and the search alike: 0x13 begins a field, 0x14 separates its code
/// from its result, 0x15 ends it (LibreOffice ww8scan.cxx: the PlcFld marks, ch = data byte &amp; 0x1f, are these
/// characters); everything goes through the shared <see cref="WordFieldStack"/> (codes hidden, a nested field's result part
/// of its parent's code, page placeholders, hidden fields, computed results). The field table gives each field's dirty /
/// locked flags and the fields open where a block starts; a 0x01 inside a FORMCHECKBOX / FORMDROPDOWN / FORMTEXT code points
/// at its FFData (read by the caller's <see cref="FormData"/>).
/// </summary>
internal sealed class DocBinaryFieldReader
{
    /// <summary>What a character does.</summary>
    public readonly record struct Step(bool IsConsumed, WordFieldOutput Output, string? LinkOpen, bool LinkClose, string? OutputLink);

    private readonly WordFieldStack _stack;
    private readonly DocBinaryFields _table;
    private readonly Func<int, WordFieldSequenceCounters>? _countersBefore;
    private readonly List<int> _beginCps = new();      // per open field: its begin CP
    private int _linkDepth = -1;                        // the stack depth of the field whose link is open (-1: none)
    private bool _countersLoaded;

    public Func<uint, WordFormFieldState?>? FormData { get; init; }

    /// <summary>The data location of a character (sprmCPicLocation, also under sprmCFData): asked for a 0x01 in a FORM* code only.</summary>
    public Func<int, uint?>? DataLocationAt { get; init; }

    /// <param name="countersBefore">The SEQ / AUTONUM / LISTNUM counters at a CP (main text only); null: counted fields keep their results.</param>
    public DocBinaryFieldReader(DocBinaryFields table, WordFieldContext? context, Func<int, WordFieldSequenceCounters>? countersBefore)
    {
        _table = table;
        _countersBefore = countersBefore;
        _stack = new WordFieldStack(context) { LinkTagOf = DocxHtml.DocxBlockHtmlRenderer.HyperlinkFieldTag };
    }

    public int Depth => _stack.Depth;

    /// <summary>Inside a SHAPE field's result (LibreOffice ww8par.cxx:3594-3622: there a lone 0x01 is not a picture of its own).</summary>
    public bool IsInShapeResult => _stack.Depth > 0 && _stack.Fields[_stack.Depth - 1] is { IsSeparated: true, Parsed.Name: "SHAPE" };

    /// <summary>
    /// The fields open at a block's start, rebuilt from the field table: each open field's code is read again through this
    /// reader (outputs dropped), then its separator if it is before <paramref name="cp"/>.
    /// </summary>
    public void RestoreAt(int cp, Func<int, int, string> readText)
    {
        List<(int beginCp, int separatorCp)> open = _table.OpenAt(cp);
        int replayedTo = -1;
        foreach ((int beginCp, int separatorCp) in open)
        {
            if (beginCp < replayedTo) break;                    // inside a code already read: the replay opened it
            int codeEnd = separatorCp >= 0 && separatorCp < cp ? separatorCp : cp;
            Consume('\u0013', beginCp, null);
            string code = readText(beginCp + 1, codeEnd);
            for (int i = 0; i < code.Length; i++) Consume(code[i], beginCp + 1 + i, null, quiet: true);
            replayedTo = codeEnd;
            if (separatorCp >= 0 && separatorCp < cp) { Consume('\u0014', separatorCp, null, quiet: true); replayedTo = separatorCp + 1; }
        }
        _linkDepth = -1;                                       // the caller reopens the active link (ReopenLink)
    }

    /// <summary>At a paragraph's start: the link of a link field still open, to reopen (null: none).</summary>
    public string? ReopenLink()
    {
        if (_linkDepth >= 0 || _stack.ActiveLink is not { } link) return null;
        for (int i = _stack.Depth - 1; i >= 0; i--)
            if (_stack.Fields[i].LinkTag == link && _stack.Fields[i].IsSeparated) { _linkDepth = i; break; }
        return link;
    }

    /// <summary>At a paragraph's end: whether a link is open (the caller closes it; <see cref="ReopenLink"/> opens it again).</summary>
    public bool CloseLinkAtParagraphEnd()
    {
        if (_linkDepth < 0) return false;
        _linkDepth = -1;
        return true;
    }

    public Step Consume(char c, int cp, uint? picture, bool quiet = false)
    {
        switch (c)
        {
            case '\u0013':
            {
                (bool dirty, bool locked, bool isPrivate) = _table.FlagsAt(cp);
                _stack.Begin(dirty, locked);
                if (isPrivate) _stack.Fields[_stack.Depth - 1].IsPrivateResult = true;
                _beginCps.Add(cp);
                return new(true, WordFieldOutput.None, null, false, null);
            }
            case '\u0014':
            {
                if (_stack.Depth == 0 || _stack.Fields[_stack.Depth - 1].IsSeparated) return new(true, WordFieldOutput.None, null, false, null);
                LoadCountersFor(_stack.Depth - 1);
                WordFieldOutput output = _stack.Separate();
                if (quiet) return new(true, WordFieldOutput.None, null, false, null);
                bool shows = _stack.OutputShows;
                string? linkOpen = null;
                WordFieldStack.Field field = _stack.Fields[_stack.Depth - 1];
                if (shows && field.LinkTag is { } tag && _linkDepth < 0 && !field.SuppressesResult) { _linkDepth = _stack.Depth - 1; linkOpen = tag; }
                return new(true, shows ? output : WordFieldOutput.None, linkOpen, false, shows && _linkDepth < 0 ? field.LinkTag : null);
            }
            case '\u0015':
            {
                if (_stack.Depth == 0) return new(true, WordFieldOutput.None, null, false, null);
                if (!_stack.Fields[_stack.Depth - 1].IsSeparated) LoadCountersFor(_stack.Depth - 1);
                bool closesLink = _linkDepth == _stack.Depth - 1;
                if (closesLink) _linkDepth = -1;
                WordFieldOutput output = _stack.End(out string? linkTag);
                _beginCps.RemoveAt(_beginCps.Count - 1);
                bool shows = _stack.Route() == WordFieldTextRoute.Visible;
                return new(true, shows && !quiet ? output : WordFieldOutput.None, null, closesLink && !quiet, shows && _linkDepth < 0 ? linkTag : null);
            }
        }
        if (_stack.Depth == 0) return new(false, WordFieldOutput.None, null, false, null);
        // a 0x01 in a form field's code: its FFData
        if (c == '\u0001' && FormData is not null && _stack.Fields[_stack.Depth - 1] is { IsSeparated: false, Form: null } owner
            && owner.Instruction.ToString().TrimStart().StartsWith("FORM", StringComparison.OrdinalIgnoreCase)
            && DataLocationAt?.Invoke(cp) is { } location)
            owner.Form = FormData(location);
        // control characters are not instruction text (HYPERLINK codes carry a 0x01 object mark: Word's hyperlink data)
        WordFieldTextRoute route = _stack.Route(c < ' ' ? " " : c.ToString());
        return new(route != WordFieldTextRoute.Visible, WordFieldOutput.None, null, false, null);
    }

    /// <summary>Before a counted field computes: the counters as they stand at its begin (once per reader; it counts on itself).</summary>
    private void LoadCountersFor(int depth)
    {
        if (_countersLoaded || _countersBefore is null) return;
        string name = WordFieldInstruction.Parse(_stack.Fields[depth].Instruction.ToString()).Name;
        if (!WordFieldEvaluator.IsCounted(name)) return;
        _stack.Counters = _countersBefore(_beginCps[depth]);
        _countersLoaded = true;
    }
}

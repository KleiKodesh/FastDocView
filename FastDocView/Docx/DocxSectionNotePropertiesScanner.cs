using SeekableDeflate;
using System.Globalization;
using System.Text;

namespace FastDocView.Docx;

/// <summary>
/// A first open's up-front pass for note numbering: a section's w:footnotePr / w:endnotePr sit in its w:sectPr at the END of
/// the section, so the windowed index learns them only once it scans past the section - too late for the notes the reader
/// already sees. This pass streams document.xml once, tag names only (no text, no DOM, a few bytes of state), and records each
/// section's end offset and its rules, so every note shows its printed number from the first window. It runs only when the
/// package has notes. The section breaks it accepts are the index scanner's (<see cref="DocxBodyBlockScanner"/>): a w:sectPr
/// that is a child of w:body, or of a w:pPr outside tables (a w:sectPrChange's older w:sectPr is not: its parent is w:sectPrChange).
/// </summary>
internal static class DocxSectionNotePropertiesScanner
{
    /// <summary>Per section, in order: the uncompressed offset just past its w:sectPr, and its footnote / endnote rules.</summary>
    public sealed record Result(long[] SectionEndOffsets, WordDocumentNoteNumberRule[] Footnotes, WordDocumentNoteNumberRule[] Endnotes);

    private enum Kind : byte { Other, Body, Table, ParagraphProperties, SectionProperties, FootnoteProperties, EndnoteProperties, NumberFormat, NumberStart, NumberRestart }

    public static Result Scan(Stream package, ZipEntryLocation entry)
    {
        var cursor = new DocxEntryInflationCursor(package, entry, long.MaxValue);
        var state = new State();
        long offset = 0;
        int read;
        while ((read = cursor.ReadNextChunk()) > 0)
        {
            state.Feed(cursor.Chunk, read, offset);
            offset += read;
        }
        return new Result(state.Ends.ToArray(), state.Footnotes.ToArray(), state.Endnotes.ToArray());
    }

    private sealed class State
    {
        public readonly List<long> Ends = new();
        public readonly List<WordDocumentNoteNumberRule> Footnotes = new(), Endnotes = new();

        private Kind[] _stack = new Kind[64];                // open elements (kinds only), [0, _depth)
        private int _depth;
        private readonly byte[] _tag = new byte[512];        // the current tag's bytes after '<' (a longer tag keeps its first 512)
        private int _tagLength;
        private bool _isInTag;
        private byte _lastTagByte;                           // the tag's last byte before '>' (a long tag keeps only its first 512)
        private int _openTables;
        private int _sectionDepth = -1;                      // stack depth of the accepted w:sectPr being read, -1 outside
        private int _noteKind = -1;                          // 0 footnotePr, 1 endnotePr inside it
        private readonly WordDocumentNoteNumberRule[] _rules = new WordDocumentNoteNumberRule[2];

        /// <summary>Vectorized: jumps '<' to '>' with IndexOf and looks only at tag names (a byte loop over 87 MB cost 3.7 s; this
        /// runs at about the inflater's speed). A tag cut by the chunk end continues into the next chunk.</summary>
        public void Feed(byte[] chunk, int length, long chunkOffset)
        {
            ReadOnlySpan<byte> data = chunk.AsSpan(0, length);
            int i = 0;
            if (_isInTag)
            {
                int close = data.IndexOf((byte)'>');
                if (close < 0) { Append(data); return; }
                Append(data.Slice(0, close));
                _isInTag = false;
                OnTag(chunkOffset + close + 1);
                i = close + 1;
            }
            while (i < length)
            {
                int open = data.Slice(i).IndexOf((byte)'<');
                if (open < 0) return;
                int start = i + open + 1;
                _tagLength = 0; _lastTagByte = 0;
                int close = data.Slice(start).IndexOf((byte)'>');
                if (close < 0) { _isInTag = true; Append(data.Slice(start)); return; }
                Append(data.Slice(start, close));
                OnTag(chunkOffset + start + close + 1);
                i = start + close + 1;
            }
        }

        private void Append(ReadOnlySpan<byte> bytes)
        {
            if (bytes.IsEmpty) return;
            _lastTagByte = bytes[bytes.Length - 1];
            int room = Math.Min(bytes.Length, _tag.Length - _tagLength);
            bytes.Slice(0, room).CopyTo(_tag.AsSpan(_tagLength));
            _tagLength += room;
        }

        private void OnTag(long endOffset)
        {
            if (_tagLength == 0 || _tag[0] is (byte)'?' or (byte)'!') return;      // declaration, comment, CDATA
            bool isClosing = _tag[0] == (byte)'/';
            bool isSelfClosing = !isClosing && _lastTagByte == (byte)'/';
            if (isClosing)
            {
                if (_depth == 0) return;
                Kind closed = _stack[--_depth];
                if (closed == Kind.Table && _openTables > 0) _openTables--;
                if (closed is Kind.FootnoteProperties or Kind.EndnoteProperties && _depth == _sectionDepth + 1) _noteKind = -1;
                if (closed == Kind.SectionProperties && _depth == _sectionDepth) CloseSection(endOffset);
                return;
            }
            int nameEnd = 0;
            while (nameEnd < _tagLength && _tag[nameEnd] is not ((byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n' or (byte)'/')) nameEnd++;
            Kind kind = Classify(_tag.AsSpan(0, nameEnd));
            Kind parent = _depth > 0 ? _stack[_depth - 1] : Kind.Other;
            int depth = _depth;                               // the new element's depth
            if (kind == Kind.SectionProperties && _sectionDepth < 0
                && (parent == Kind.Body || (parent == Kind.ParagraphProperties && _openTables == 0)))
            {
                _sectionDepth = depth;
                _noteKind = -1;
                _rules[0] = WordDocumentNoteNumberRule.FootnoteDefault;
                _rules[1] = WordDocumentNoteNumberRule.EndnoteDefault;
                if (isSelfClosing) { CloseSection(endOffset); return; }
            }
            else if (_sectionDepth >= 0 && depth == _sectionDepth + 1 && kind is Kind.FootnoteProperties or Kind.EndnoteProperties)
                _noteKind = isSelfClosing ? -1 : kind == Kind.FootnoteProperties ? 0 : 1;
            else if (_sectionDepth >= 0 && depth == _sectionDepth + 2 && _noteKind >= 0 && kind is Kind.NumberFormat or Kind.NumberStart or Kind.NumberRestart)
                ReadRule(kind);
            if (kind == Kind.Table) _openTables++;
            if (!isSelfClosing)
            {
                if (_depth == _stack.Length) Array.Resize(ref _stack, _stack.Length * 2);
                _stack[_depth++] = kind;
            }
            else if (kind == Kind.Table) _openTables--;
        }

        private void CloseSection(long endOffset)
        {
            _sectionDepth = -1;
            _noteKind = -1;
            Ends.Add(endOffset);
            Footnotes.Add(_rules[0]);
            Endnotes.Add(_rules[1]);
        }

        private void ReadRule(Kind kind)
        {
            string? value = Value("w:val=");
            if (value is null) return;
            WordDocumentNoteNumberRule rule = _rules[_noteKind];
            if (kind == Kind.NumberFormat) rule = rule with { Format = value };
            else if (kind == Kind.NumberStart) { if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int start) && start >= 0) rule = rule with { Start = start }; }
            else rule = rule with { Restart = value switch { "eachSect" => WordDocumentNoteRestart.EachSection, "eachPage" => WordDocumentNoteRestart.EachPage, _ => WordDocumentNoteRestart.Continuous } };
            _rules[_noteKind] = rule;
        }

        private string? Value(string attributeWithEquals)
        {
            string tag = Encoding.UTF8.GetString(_tag, 0, _tagLength);
            int at = tag.IndexOf(" " + attributeWithEquals, StringComparison.Ordinal);
            if (at < 0) return null;
            at += attributeWithEquals.Length + 1;
            if (at >= tag.Length) return null;
            char quote = tag[at];
            int end = tag.IndexOf(quote, at + 1);
            return end > at ? tag.Substring(at + 1, end - at - 1) : null;
        }
    }

    private static Kind Classify(ReadOnlySpan<byte> name) => name.Length switch
    {
        5 when name.SequenceEqual("w:tbl"u8) => Kind.Table,
        5 when name.SequenceEqual("w:pPr"u8) => Kind.ParagraphProperties,
        6 when name.SequenceEqual("w:body"u8) => Kind.Body,
        8 when name.SequenceEqual("w:sectPr"u8) => Kind.SectionProperties,
        8 when name.SequenceEqual("w:numFmt"u8) => Kind.NumberFormat,
        10 when name.SequenceEqual("w:numStart"u8) => Kind.NumberStart,
        11 when name.SequenceEqual("w:endnotePr"u8) => Kind.EndnoteProperties,
        12 when name.SequenceEqual("w:footnotePr"u8) => Kind.FootnoteProperties,
        12 when name.SequenceEqual("w:numRestart"u8) => Kind.NumberRestart,
        _ => Kind.Other,
    };
}

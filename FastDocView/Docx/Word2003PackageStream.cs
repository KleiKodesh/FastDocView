using System.Text;
using System.Xml;
using SeekableDeflate;

namespace FastDocView.Docx;

/// <summary>
/// A Word 2003 XML file (w:wordDocument) presented as a .docx package, windowed like every other format: nothing is
/// translated whole. At open, one byte pass (<see cref="Word2003SourceScanner"/>: no XML parse, no DOM) finds the body's
/// byte range, every picture's base64 range (w:binData), every hyperlink target (w:hlink w:dest) and whether notes
/// exist; the header parts before the body (w:styles, w:lists, w:docPr: small) are translated into styles.xml,
/// numbering.xml and settings.xml, and the relationships are made from what the pass found. The body and the notes are
/// derived parts (<see cref="DocxVirtualPackageStream.DerivedPartMethod"/>): the .docx reader's index pass pulls them
/// block by block (each body child is cut out by <see cref="Word2003BodyBlockSplitter"/> and translated on its own by
/// <see cref="Word2003XmlTranslator.Translation"/>), recording a checkpoint (the block's byte offset in the file, its
/// offset in the translated part) about every <see cref="TranslationCheckpointSpanBytes"/>; a window reads from the
/// nearest checkpoint, translating only the blocks up to it. The checkpoints live in the sidecar index like deflate
/// checkpoints (with no window bytes). A block's translation depends only on its own bytes (note ids come from its
/// offset, relationship ids from the open pass), so a block translated again is byte-identical to the index pass's.
/// </summary>
internal sealed class Word2003PackageStream : DocxVirtualPackageStream, IDocxDerivedPartSource
{
    /// <summary>
    /// Translated bytes between checkpoints. Denser than deflate's 2 MB: a translation checkpoint costs two numbers in the
    /// index, while reaching a block past one costs translating everything in between.
    /// </summary>
    public const long TranslationCheckpointSpanBytes = 64 << 10;

    private const string DocumentPartName = "word/document.xml", FootnotesPartName = "word/footnotes.xml", EndnotesPartName = "word/endnotes.xml";

    private readonly Word2003SourceLayout _layout;
    private readonly Word2003XmlTranslator.TranslationState _state = new();
    private readonly Dictionary<Word2003XmlTranslator.NotePass, (byte[] Prefix, byte[] Suffix)> _frames = new();

    // The last file's open pass, keyed by its path, length and write time: a host asks Decide and then Open for the same
    // file, and the second needs no second pass over it. One entry (bounded); the layout is immutable.
    private static readonly object LastScanGate = new();
    private static (string Key, Word2003SourceLayout Layout)? _lastScan;
    private static int _scanCount;

    /// <summary>Open passes made over Word 2003 XML files so far (the harness checks that Decide + Open make one).</summary>
    internal static int ScanCount => Volatile.Read(ref _scanCount);

    /// <summary>Forgets the last open pass (the harness times a cold open).</summary>
    internal static void ForgetLastScan() { lock (LastScanGate) _lastScan = null; }

    /// <summary>The package over <paramref name="file"/> (owned). With <paramref name="filePath"/>, the open pass of an unchanged file is reused.</summary>
    public Word2003PackageStream(Stream file, string? filePath = null) : base(file)
    {
        string? key = null;
        if (filePath is not null)
        {
            var info = new FileInfo(filePath);
            key = Path.GetFullPath(filePath) + "|" + info.Length + "|" + info.LastWriteTimeUtc.Ticks;
        }
        Word2003SourceLayout? layout = null;
        lock (LastScanGate) if (key is not null && _lastScan is { } last && last.Key == key) layout = last.Layout;
        if (layout is null)
        {
            DocxFlatPackageStream.RequireUtf8(file, "Word 2003 XML files");
            layout = Word2003SourceScanner.Scan(file);
            Interlocked.Increment(ref _scanCount);
            if (key is not null) lock (LastScanGate) _lastScan = (key, layout);
        }
        _layout = layout;

        foreach (string target in layout.HyperlinkTargets) _state.HyperlinkId(target);
        var pictures = new List<Word2003SourceLayout.Picture>();
        foreach (Word2003SourceLayout.Picture picture in layout.Pictures)
            if (!_state.ImageRelationships.ContainsKey(picture.Name)) { _state.ImageId(picture.Name); pictures.Add(picture); }   // the first of a repeated name
        _state.Freeze();

        Word2003XmlTranslator.Header header = Word2003XmlTranslator.ReadHeader(file);
        var parts = new List<VirtualPart>();
        foreach ((string name, byte[] bytes) in Word2003XmlTranslator.SynthesizedParts(header, _state, layout.HasFootnotes, layout.HasEndnotes))
            parts.Add(VirtualPart.Synthesized(name, bytes));
        // the body's own bytes stand for the body part's size where the reader weighs parts (the notes are inside them)
        parts.Add(VirtualPart.Derived(DocumentPartName, Math.Max(0, layout.BodyEnd - layout.BodyStart)));
        if (layout.HasFootnotes) parts.Add(VirtualPart.Derived(FootnotesPartName, 0));
        if (layout.HasEndnotes) parts.Add(VirtualPart.Derived(EndnotesPartName, 0));
        foreach (Word2003SourceLayout.Picture picture in pictures)
            parts.Add(VirtualPart.Base64(Word2003XmlTranslator.MediaPartPath(_state, picture.Name), picture.TextStart, picture.TextEnd, picture.DecodedLength));
        Layout(parts);
    }

    private static Word2003XmlTranslator.NotePass PassOf(ZipEntryLocation entry) => entry.Name switch
    {
        FootnotesPartName => Word2003XmlTranslator.NotePass.Footnotes,
        EndnotesPartName => Word2003XmlTranslator.NotePass.Endnotes,
        DocumentPartName => Word2003XmlTranslator.NotePass.None,
        _ => throw new InvalidDataException($"Word 2003 XML: no derived part {entry.Name}."),
    };

    public IDocxDerivedPartPass OpenPass(ZipEntryLocation entry, List<DeflateCheckpoint> checkpoints, long checkpointSpanBytes)
        => new TranslationPass(this, PassOf(entry), checkpoints);

    public Stream OpenView(ZipEntryLocation entry, IReadOnlyList<DeflateCheckpoint> checkpoints)
        => new TranslatedPartView(this, PassOf(entry), checkpoints);

    /// <summary>The bytes before the first block (XML declaration, root start tag, w:body start tag) and after the last.</summary>
    private (byte[] Prefix, byte[] Suffix) FramesOf(Word2003XmlTranslator.NotePass pass)
    {
        if (_frames.TryGetValue(pass, out var frames)) return frames;
        var run = new TranslationRun(this, pass, _layout.BodyStart);
        frames = (run.Prefix, run.Suffix());
        _frames[pass] = frames;
        return frames;
    }

    // ---- one translation run: blocks from a source offset, in order ----------------------------------------------------

    /// <summary>
    /// Translates the body's blocks one after another from a block start. One XmlWriter carries the root element's
    /// namespace scope (its start tag is the part's prefix), so each block comes out exactly as in a whole-part translation.
    /// </summary>
    private sealed class TranslationRun
    {
        private static readonly XmlReaderSettings BlockReaderSettings = new()
        {
            ConformanceLevel = ConformanceLevel.Fragment, DtdProcessing = DtdProcessing.Ignore, CloseInput = true,
            IgnoreComments = true, IgnoreProcessingInstructions = true,
        };

        private readonly Word2003PackageStream _owner;
        private readonly Word2003XmlTranslator.NotePass _pass;
        private readonly Word2003BodyBlockSplitter _splitter;
        private readonly MemoryStream _output = new();
        private readonly XmlWriter _writer;
        private readonly Word2003XmlTranslator.Translation _translation;
        private readonly byte[] _noteTag;
        private byte[] _block = new byte[1 << 16];
        private bool _isCut;                  // a malformed block ended the part: nothing after it is read

        public byte[] Prefix { get; }
        public bool IsFinished { get; private set; }

        /// <summary>The file offset of the next block to translate (where a checkpoint taken now points).</summary>
        public long NextBlockOffset => _splitter.Position;

        public TranslationRun(Word2003PackageStream owner, Word2003XmlTranslator.NotePass pass, long sourceOffset)
        {
            _owner = owner;
            _pass = pass;
            _splitter = new Word2003BodyBlockSplitter(owner.SourceFile, sourceOffset, owner._layout);
            _writer = XmlWriter.Create(_output, new XmlWriterSettings { Encoding = new UTF8Encoding(false), CloseOutput = false });
            _translation = new Word2003XmlTranslator.Translation(_writer, owner._state, pass);
            _writer.WriteStartDocument(standalone: true);
            _translation.StartRoot(pass switch
            {
                Word2003XmlTranslator.NotePass.Footnotes => "footnotes",
                Word2003XmlTranslator.NotePass.Endnotes => "endnotes",
                _ => "document",
            });
            if (pass == Word2003XmlTranslator.NotePass.None) _writer.WriteStartElement("w", "body", Word2003XmlTranslator.DocxMainNamespace);
            _writer.WriteString("");          // closes the start tag: the next block starts on its own '<'
            Prefix = TakeOutput();
            if (Prefix.Length == 0 || Prefix[Prefix.Length - 1] != (byte)'>') throw new InvalidOperationException("Word 2003 XML: the part's start tag was not closed.");
            _noteTag = pass == Word2003XmlTranslator.NotePass.None ? []
                     : Encoding.UTF8.GetBytes("<" + owner._layout.WordQualifier + (pass == Word2003XmlTranslator.NotePass.Footnotes ? "footnote" : "endnote"));
        }

        private byte[] TakeOutput()
        {
            _writer.Flush();
            byte[] bytes = _output.ToArray();
            _output.SetLength(0);
            return bytes;
        }

        /// <summary>The next block's translated bytes (empty for a block a notes part has nothing from); null at the part's end.</summary>
        public byte[]? NextBlock()
        {
            if (IsFinished) return null;
            if (_isCut || !_splitter.Next(ref _block, out long blockStart, out int length))
            {
                IsFinished = true;
                return null;
            }
            if (_noteTag.Length > 0 && !Word2003SourceScanner.ContainsTag(_block, length, _noteTag)) return [];
            _owner._state.BeginBlock(blockStart);
            try
            {
                using XmlReader reader = XmlReader.Create(new MemoryStream(_block, 0, length, writable: false), BlockReaderSettings, _owner._layout.BlockContext());
                reader.MoveToContent();
                if (reader.NodeType == XmlNodeType.Element) _translation.Element(reader);
            }
            catch (XmlException)
            {
                // Malformed: keep what precedes the first error, as for a .docx part (WordDocumentXml.LoadUntilFirstError):
                // the elements still open are closed, the rest of the file is not read.
                _translation.CloseAll();
                _isCut = true;
            }
            return TakeOutput();
        }

        /// <summary>The part's closing tags (written by a run that translated no block).</summary>
        public byte[] Suffix()
        {
            _translation.CloseAll();
            if (_pass == Word2003XmlTranslator.NotePass.None) _writer.WriteEndElement();     // w:body
            _writer.WriteEndElement();                                                          // the root
            return TakeOutput();
        }
    }

    // ---- the index pass ------------------------------------------------------------------------------------------------

    private sealed class TranslationPass(Word2003PackageStream owner, Word2003XmlTranslator.NotePass pass, List<DeflateCheckpoint> checkpoints)
        : IDocxDerivedPartPass
    {
        private readonly TranslationRun _run = new(owner, pass, owner._layout.BodyStart);
        private byte[] _pending = [];
        private int _pendingAt;
        private long _translated;             // part offset of the next byte handed out
        private long _nextCheckpointAt;
        private int _stage;                   // 0: prefix next, 1: blocks, 2: suffix handed out

        public int Read(byte[] buffer, int offset, int count)
        {
            while (_pendingAt == _pending.Length)
            {
                if (_stage == 0) { _pending = _run.Prefix; _stage = 1; }
                else if (_stage == 1)
                {
                    long blockOffset = _run.NextBlockOffset;
                    bool isCheckpointDue = _translated >= _nextCheckpointAt;
                    byte[]? block = _run.NextBlock();
                    if (block is null) { _pending = owner.FramesOf(pass).Suffix; _stage = 2; }
                    else
                    {
                        if (isCheckpointDue && block.Length > 0)
                        {
                            checkpoints.Add(new DeflateCheckpoint(blockOffset, 0, 0, _translated, []));
                            _nextCheckpointAt = _translated + TranslationCheckpointSpanBytes;
                        }
                        _pending = block;
                    }
                }
                else return 0;
                _pendingAt = 0;
            }
            int take = Math.Min(count, _pending.Length - _pendingAt);
            Buffer.BlockCopy(_pending, _pendingAt, buffer, offset, take);
            _pendingAt += take;
            _translated += take;
            return take;
        }
    }

    // ---- random access -------------------------------------------------------------------------------------------------

    /// <summary>
    /// A seekable view of a translated part: the prefix, then blocks translated from the nearest checkpoint at or before
    /// the position read (the part's first block when there is none). The last run's output is kept (bounded), so the
    /// next window's read, usually just after, continues it instead of starting again.
    /// </summary>
    private sealed class TranslatedPartView(Word2003PackageStream owner, Word2003XmlTranslator.NotePass pass, IReadOnlyList<DeflateCheckpoint> checkpoints)
        : Stream
    {
        private const int KeptBytes = 4 << 20;

        private readonly byte[] _prefix = owner.FramesOf(pass).Prefix;
        private long _position;
        private TranslationRun? _run;
        private byte[] _kept = [];            // translated bytes [_keptStart, _keptStart + _keptLength): the run's latest output
        private long _keptStart;
        private int _keptLength;
        private long _end = -1;               // the part's length once a run reached its end
        private bool _runEnded;               // the current run appended the suffix

        public override int Read(byte[] buffer, int offset, int count)
        {
            if (count <= 0 || (_end >= 0 && _position >= _end)) return 0;
            if (_position < _prefix.Length)
            {
                int fromPrefix = (int)Math.Min(count, _prefix.Length - _position);
                Buffer.BlockCopy(_prefix, (int)_position, buffer, offset, fromPrefix);
                _position += fromPrefix;
                return fromPrefix;
            }
            if (!(_run is not null && _position >= _keptStart && _position <= _keptStart + _keptLength + TranslationCheckpointSpanBytes)) Restart();
            while (_position >= _keptStart + _keptLength)
                if (!Extend()) break;
            if (_end >= 0 && _position >= _end) return 0;
            int within = (int)(_position - _keptStart);
            int take = Math.Min(count, _keptLength - within);
            Buffer.BlockCopy(_kept, within, buffer, offset, take);
            _position += take;
            return take;
        }

        /// <summary>A new run from the last checkpoint at or before the position.</summary>
        private void Restart()
        {
            long source = owner._layout.BodyStart, translated = _prefix.Length;
            int low = 0, high = checkpoints.Count - 1;
            while (low <= high)
            {
                int middle = (low + high) / 2;
                if (checkpoints[middle].UncompressedOffset <= _position) { source = checkpoints[middle].CompressedBytePosition; translated = checkpoints[middle].UncompressedOffset; low = middle + 1; }
                else high = middle - 1;
            }
            _run = new TranslationRun(owner, pass, source);
            _keptStart = translated;
            _keptLength = 0;
            _runEnded = false;
        }

        /// <summary>Appends the run's next block (or the suffix at the end); false once the part is exhausted.</summary>
        private bool Extend()
        {
            byte[]? block = _run!.NextBlock();
            if (block is null)
            {
                if (_runEnded) return false;
                block = owner.FramesOf(pass).Suffix;
                _end = _keptStart + _keptLength + block.Length;
                _runEnded = true;
            }
            if (block.Length == 0) return true;
            // drop what lies before the position once the kept bytes would outgrow their bound
            if (_keptLength + block.Length > KeptBytes && _position > _keptStart)
            {
                int drop = (int)Math.Min(_keptLength, _position - _keptStart);
                Buffer.BlockCopy(_kept, drop, _kept, 0, _keptLength - drop);
                _keptStart += drop;
                _keptLength -= drop;
            }
            if (_keptLength + block.Length > _kept.Length) Array.Resize(ref _kept, Math.Max(_keptLength + block.Length, Math.Min(KeptBytes, Math.Max(1 << 16, _kept.Length * 2))));
            Buffer.BlockCopy(block, 0, _kept, _keptLength, block.Length);
            _keptLength += block.Length;
            return true;
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            _position = origin switch
            {
                SeekOrigin.Begin => offset,
                SeekOrigin.Current => _position + offset,
                _ => throw new NotSupportedException("A translated part's length is known only once it was read to its end."),
            };
            return _position;
        }

        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => _end >= 0 ? _end : throw new NotSupportedException("A translated part's length is known only once it was read to its end.");
        public override long Position { get => _position; set => _position = value; }
        public override void Flush() { }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}

/// <summary>What the open pass over a Word 2003 XML file found.</summary>
internal sealed class Word2003SourceLayout
{
    public sealed record Picture(string Name, long TextStart, long TextEnd, long DecodedLength);

    /// <summary>The body's content: just after the w:body start tag, up to the '&lt;' of its end tag.</summary>
    public long BodyStart, BodyEnd;
    /// <summary>"w:" (or "" for a default namespace): the prefix the root binds to the 2003 namespace.</summary>
    public string WordQualifier = "w:";
    /// <summary>"wx:": the prefix of the auxiliary namespace (wx:sect / wx:sub-section wrap the body's blocks); null when unbound.</summary>
    public string? AuxiliaryQualifier;
    public readonly List<(string Prefix, string Namespace)> RootNamespaces = new();
    public readonly List<Picture> Pictures = new();
    public readonly List<string> HyperlinkTargets = new();
    public bool HasFootnotes, HasEndnotes;

    /// <summary>A fresh parser context for one block: the root's namespace declarations and its xml:space.</summary>
    public XmlParserContext BlockContext()
    {
        var names = new NameTable();
        var namespaces = new XmlNamespaceManager(names);
        foreach ((string prefix, string ns) in RootNamespaces) namespaces.AddNamespace(prefix, ns);
        return new XmlParserContext(names, namespaces, null, XmlSpace.Preserve);
    }
}

/// <summary>
/// The open pass over a Word 2003 XML file: a byte scan (no XML parse, bounded memory) for the few tags whose
/// information the package needs before any block is translated. Only start tags in the 2003 namespace are looked at:
/// w:body (the body's range), w:binData (each picture's base64 text and decoded length), w:hlink (w:dest: one
/// relationship per target) and, inside the body, w:footnote / w:endnote (whether the notes parts exist). Text cannot
/// hold a '&lt;' (it is escaped), and Word writes no comments or CDATA in the body.
/// </summary>
internal static class Word2003SourceScanner
{
    private const int ChunkSize = 1 << 20;
    private const string RootNamespaceForAuxiliary = "http://schemas.microsoft.com/office/word/2003/auxHint";

    public static Word2003SourceLayout Scan(Stream file)
    {
        var layout = new Word2003SourceLayout();
        ReadRoot(file, layout);
        var reader = new Word2003ByteReader(file, 0, ChunkSize);
        string q = layout.WordQualifier;
        const int Body = 0, BodyEnd = 1, BinData = 2, Hyperlink = 3, Footnote = 4, Endnote = 5;
        byte[][] tags = [.. new[] { "<" + q + "body", "</" + q + "body", "<" + q + "binData", "<" + q + "hlink", "<" + q + "footnote", "<" + q + "endnote" }.Select(Encoding.UTF8.GetBytes)];
        byte[] binDataEnd = Encoding.UTF8.GetBytes("</" + q + "binData");
        var targets = new HashSet<string>(StringComparer.Ordinal);
        bool inBody = false;
        layout.BodyStart = layout.BodyEnd = -1;
        bool[] active = [true, true, true, true, true, true];
        long[] next = [long.MinValue, long.MinValue, long.MinValue, long.MinValue, long.MinValue, long.MinValue];
        long[] searchedTo = new long[tags.Length];
        while (reader.FindTag(tags, active, next, searchedTo) is int found and >= 0)
        {
            // a tag needed once is no longer looked for after it was found
            if (found == Body || found == BodyEnd) active[found] = false;
            if ((found == Footnote || found == Endnote) && inBody) active[found] = false;
            switch (found)
            {
                case Body when layout.BodyStart < 0:
                {
                    string tag = reader.ReadTag();
                    layout.BodyStart = reader.Position;
                    inBody = !tag.EndsWith("/", StringComparison.Ordinal);
                    if (!inBody) layout.BodyEnd = layout.BodyStart;
                    break;
                }
                case BodyEnd when inBody:
                    layout.BodyEnd = reader.Position;
                    inBody = false;
                    reader.Skip(1);
                    break;
                case BinData:
                {
                    string tag = reader.ReadTag();
                    string? pictureName = AttributeValue(tag, q + "name");
                    if (tag.EndsWith("/", StringComparison.Ordinal) || pictureName is null) break;
                    long start = reader.Position;
                    (long end, long characters, int padding) = reader.Base64Until(binDataEnd);
                    if (end < 0) break;
                    layout.Pictures.Add(new Word2003SourceLayout.Picture(pictureName, start, end, Math.Max(0, characters / 4 * 3 - padding)));
                    break;
                }
                case Hyperlink:
                {
                    string tag = reader.ReadTag();
                    if (AttributeValue(tag, q + "dest") is { Length: > 0 } target && targets.Add(target)) layout.HyperlinkTargets.Add(target);
                    break;
                }
                case Footnote when inBody: layout.HasFootnotes = true; reader.Skip(1); break;
                case Endnote when inBody: layout.HasEndnotes = true; reader.Skip(1); break;
                default: reader.Skip(1); break;
            }
        }
        if (layout.BodyStart < 0) { layout.BodyStart = layout.BodyEnd = reader.Position; }   // no body: an empty document
        else if (layout.BodyEnd < 0) layout.BodyEnd = reader.Position;                       // cut off inside the body: up to the file's end
        return layout;
    }

    /// <summary>The root's namespace declarations; throws when the root is not the 2003 w:wordDocument.</summary>
    private static void ReadRoot(Stream file, Word2003SourceLayout layout)
    {
        file.Position = 0;
        using XmlReader reader = XmlReader.Create(file, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore, CloseInput = false });
        if (reader.MoveToContent() != XmlNodeType.Element || reader.LocalName != "wordDocument" || reader.NamespaceURI != Word2003XmlTranslator.WordNamespace)
            throw new InvalidDataException("Word 2003 XML: the root is not w:wordDocument.");
        layout.WordQualifier = reader.Prefix.Length == 0 ? "" : reader.Prefix + ":";
        if (reader.MoveToFirstAttribute())
            do
            {
                if (reader.NamespaceURI != "http://www.w3.org/2000/xmlns/") continue;
                string prefix = reader.Prefix == "xmlns" ? reader.LocalName : "";
                layout.RootNamespaces.Add((prefix, reader.Value));
                if (reader.Value == RootNamespaceForAuxiliary && prefix.Length > 0) layout.AuxiliaryQualifier = prefix + ":";
            }
            while (reader.MoveToNextAttribute());
    }

    private static bool IsNameEnd(int b) => b is '>' or '/' or ' ' or '\t' or '\r' or '\n' or -1;

    /// <summary>True when <paramref name="tag"/> ("&lt;w:footnote") starts a tag somewhere in the bytes (followed by a name end).</summary>
    public static bool ContainsTag(byte[] bytes, int length, byte[] tag)
    {
        var span = new ReadOnlySpan<byte>(bytes, 0, length);
        int from = 0;
        while (true)
        {
            int at = span.Slice(from).IndexOf(tag);
            if (at < 0) return false;
            int after = from + at + tag.Length;
            if (after >= length || IsNameEnd(bytes[after])) return true;
            from = after;
        }
    }

    /// <summary>An attribute's value in a start tag's text (entities decoded, white space normalized as an XML parser does).</summary>
    internal static string? AttributeValue(string tag, string attribute)
    {
        int at = 0;
        while ((at = tag.IndexOf(attribute, at, StringComparison.Ordinal)) >= 0)
        {
            int equals = at + attribute.Length;
            while (equals < tag.Length && char.IsWhiteSpace(tag[equals])) equals++;
            if (at > 0 && char.IsWhiteSpace(tag[at - 1]) && equals < tag.Length && tag[equals] == '=')
            {
                int quote = equals + 1;
                while (quote < tag.Length && char.IsWhiteSpace(tag[quote])) quote++;
                if (quote >= tag.Length) return null;
                char q = tag[quote];
                int close = tag.IndexOf(q, quote + 1);
                if (close < 0) return null;
                string raw = tag.Substring(quote + 1, close - quote - 1).Replace('\t', ' ').Replace('\n', ' ').Replace('\r', ' ');   // XML 1.0 3.3.3
                return System.Net.WebUtility.HtmlDecode(raw);
            }
            at += attribute.Length;
        }
        return null;
    }
}

/// <summary>A forward byte reader over a file region in large chunks, re-seating the shared stream on every refill.</summary>
internal sealed class Word2003ByteReader
{
    private readonly Stream _file;
    private readonly byte[] _chunk;
    private long _chunkStart;
    private int _chunkLength, _at;

    public Word2003ByteReader(Stream file, long start, int chunkSize)
    {
        _file = file;
        _chunk = new byte[chunkSize];
        _chunkStart = start;
    }

    public long Position => _chunkStart + _at;

    /// <summary>Makes at least <paramref name="needed"/> bytes available from the position (fewer at the file's end).</summary>
    private int Available(int needed)
    {
        if (_chunkLength - _at >= needed) return _chunkLength - _at;
        int kept = _chunkLength - _at;
        Buffer.BlockCopy(_chunk, _at, _chunk, 0, kept);
        _chunkStart += _at;
        _at = 0;
        _file.Position = _chunkStart + kept;
        int read;
        while (kept < _chunk.Length && (read = _file.Read(_chunk, kept, _chunk.Length - kept)) > 0) kept += read;
        _chunkLength = kept;
        return kept;
    }

    /// <summary>
    /// Moves to the next occurrence of one of the active <paramref name="tags"/> ("&lt;w:binData", ...) followed by a name
    /// end; the index of the tag found (the position stays on its '&lt;'), -1 at the end. Each tag is looked for with a
    /// vectorized sequence search over the chunk and its next occurrence kept (<paramref name="next"/>,
    /// <paramref name="searchedTo"/>: absolute offsets, one per tag), so the common tags in between cost nothing.
    /// </summary>
    public int FindTag(byte[][] tags, bool[] active, long[] next, long[] searchedTo)
    {
        int longest = 0;
        foreach (byte[] tag in tags) longest = Math.Max(longest, tag.Length);
        while (true)
        {
            if (Available(longest + 1) == 0) return -1;
            long chunkEnd = _chunkStart + _chunkLength;
            int best = -1;
            long bestAt = long.MaxValue;
            for (int i = 0; i < tags.Length; i++)
            {
                if (!active[i]) continue;
                byte[] tag = tags[i];
                if (next[i] < Position || (next[i] == long.MaxValue && searchedTo[i] < chunkEnd))
                {
                    // a passed occurrence: search on from the position; none so far: from where the last search stopped
                    long from = next[i] == long.MaxValue ? Math.Max(Position, searchedTo[i] - (tag.Length - 1)) : Position;
                    int found = new ReadOnlySpan<byte>(_chunk, (int)(from - _chunkStart), (int)(chunkEnd - from)).IndexOf(tag);
                    next[i] = found < 0 ? long.MaxValue : from + found;
                    searchedTo[i] = chunkEnd;
                }
                if (next[i] < bestAt) { bestAt = next[i]; best = i; }
            }
            if (best < 0)
            {
                // nothing in this chunk: keep its last bytes (a tag may straddle the refill) and read on
                int before = _chunkLength - _at;
                _at = Math.Max(_at, _chunkLength - longest);
                if (Available(longest + 1) <= Math.Min(before, longest)) return -1;
                continue;
            }
            _at = (int)(bestAt - _chunkStart);
            int length = tags[best].Length;
            int after = PeekAt(length);
            if (after is '>' or '/' or ' ' or 9 or 10 or 13) return best;
            _at++;                                                      // a longer name ("<w:bodyX"): look further
        }
    }

    /// <summary>Moves to the next occurrence of <paramref name="value"/> (the position stays on it); false at the end.</summary>
    public bool FindByte(byte value)
    {
        while (true)
        {
            int at = new ReadOnlySpan<byte>(_chunk, _at, _chunkLength - _at).IndexOf(value);
            if (at >= 0) { _at += at; return true; }
            _at = _chunkLength;
            if (Available(1) == 0) return false;
        }
    }

    public void Skip(int count)
    {
        while (count > 0)
        {
            int available = Available(1);
            if (available == 0) return;
            int step = Math.Min(count, available);
            _at += step;
            count -= step;
        }
    }

    public bool Matches(byte[] bytes)
    {
        if (Available(bytes.Length) < bytes.Length) return false;
        return new ReadOnlySpan<byte>(_chunk, _at, bytes.Length).SequenceEqual(bytes);
    }

    /// <summary>The byte <paramref name="offset"/> past the position, -1 past the end.</summary>
    public int PeekAt(int offset) => Available(offset + 1) > offset ? _chunk[_at + offset] : -1;

    /// <summary>Reads the start tag at the position up to and past its closing '&gt;' (quotes respected); its text without the brackets.</summary>
    public string ReadTag()
    {
        var bytes = new List<byte>(128);
        byte quote = 0;
        Skip(1);                                   // '<'
        while (true)
        {
            if (Available(1) == 0) break;
            byte b = _chunk[_at++];
            if (quote != 0) { if (b == quote) quote = 0; }
            else if (b is (byte)'"' or (byte)'\'') quote = b;
            else if (b == '>') break;
            bytes.Add(b);
        }
        return Encoding.UTF8.GetString(bytes.ToArray());
    }

    /// <summary>Moves past the tag at the position (its '&lt;') to after its '&gt;' (quotes respected); true when it was self-closing.</summary>
    public bool SkipTag()
    {
        _at++;
        byte quote = 0, previous = 0;
        while (true)
        {
            if (_at >= _chunkLength && Available(1) == 0) return false;
            var span = new ReadOnlySpan<byte>(_chunk, _at, _chunkLength - _at);
            if (quote != 0)
            {
                int close = span.IndexOf(quote);
                if (close < 0) { _at = _chunkLength; continue; }
                _at += close + 1;
                quote = 0;
                previous = 0;
                continue;
            }
            int stop = span.IndexOfAny((byte)'>', (byte)'"', (byte)0x27);
            if (stop < 0) { previous = span[span.Length - 1]; _at = _chunkLength; continue; }
            byte b = span[stop];
            if (stop > 0) previous = span[stop - 1];
            _at += stop + 1;
            if (b == '>') return previous == '/';
            quote = b;
        }
    }

    /// <summary>Base64 text up to <paramref name="closing"/>: the closing tag's position (-1 when missing), the base64 character count and the '=' padding. The position ends after the closing tag's name.</summary>
    public (long end, long characters, int padding) Base64Until(byte[] closing)
    {
        long characters = 0;
        int padding = 0;
        while (true)
        {
            if (Available(1) == 0) return (-1, 0, 0);
            int length = _chunkLength - _at;
            var span = new ReadOnlySpan<byte>(_chunk, _at, length);
            int lt = span.IndexOf((byte)'<');
            int count = lt < 0 ? length : lt;
            for (int i = 0; i < count; i++)
            {
                byte b = span[i];
                if (b > (byte)' ') { characters++; if (b == '=') padding++; }
            }
            _at += count;
            if (lt < 0) continue;
            if (Matches(closing)) { long end = Position; Skip(closing.Length); return (end, characters, padding); }
            Skip(1);
        }
    }
}

/// <summary>
/// Cuts a Word 2003 body into its blocks from a block start, byte by byte: each element at the body's level is one
/// block (a paragraph, a table, a w:sectPr, a body-level annotation), and wx:sect / wx:sub-section, which only wrap the
/// body (LibreOffice wordml2ooo.xsl:244-250), are looked through. Comments, processing instructions and CDATA are
/// skipped whole, so a '&lt;' inside them is never taken for a tag.
/// </summary>
internal sealed class Word2003BodyBlockSplitter
{
    private readonly Stream _file;
    private readonly Word2003ByteReader _reader;
    private readonly long _end;
    private readonly byte[]? _sectionOpen, _subSectionOpen;

    public Word2003BodyBlockSplitter(Stream file, long start, Word2003SourceLayout layout)
    {
        _file = file;
        _reader = new Word2003ByteReader(file, start, 1 << 16);
        _end = layout.BodyEnd;
        if (layout.AuxiliaryQualifier is { } aux)
        {
            _sectionOpen = Encoding.UTF8.GetBytes("<" + aux + "sect");
            _subSectionOpen = Encoding.UTF8.GetBytes("<" + aux + "sub-section");
        }
    }

    /// <summary>The file offset the next block search starts from (a block start, or the white space before one).</summary>
    public long Position => _reader.Position;

    /// <summary>The next block's bytes into <paramref name="block"/> (grown as needed); false at the body's end.</summary>
    public bool Next(ref byte[] block, out long start, out int length)
    {
        start = -1;
        length = 0;
        while (_reader.Position < _end && _reader.FindByte((byte)'<') && _reader.Position < _end)
        {
            int kind = TagKind();
            if (kind == TagSkipped) continue;
            if (kind == TagEnd) { SkipTag(); continue; }                       // a wrapper's end tag
            if (IsWrapper()) { SkipTag(); continue; }
            start = _reader.Position;
            int depth = 0;
            while (true)
            {
                int tag = TagKind();
                if (tag == TagSkipped) { }
                else if (tag == TagEnd) { SkipTag(); depth--; }
                else { if (!SkipTag()) depth++; }
                if (depth <= 0 || _reader.Position >= _end) break;
                if (!_reader.FindByte((byte)'<') || _reader.Position >= _end) break;
            }
            long end = Math.Min(_reader.Position, _end);
            length = checked((int)(end - start));
            if (block.Length < length) Array.Resize(ref block, Math.Max(length, block.Length * 2));
            ReadExactly(start, block, length);
            return true;
        }
        return false;
    }

    private void ReadExactly(long offset, byte[] buffer, int length)
    {
        _file.Position = offset;
        int filled = 0;
        while (filled < length)
        {
            int read = _file.Read(buffer, filled, length - filled);
            if (read <= 0) throw new EndOfStreamException("The Word 2003 XML file ended inside a block.");
            filled += read;
        }
    }

    private const int TagStart = 0, TagEnd = 1, TagSkipped = 2;

    /// <summary>At a '&lt;': an end tag, a start tag, or (a comment, PI, CDATA, declaration) skipped whole here.</summary>
    private int TagKind()
    {
        int next = _reader.PeekAt(1);
        if (next == '/') return TagEnd;
        if (next == '?') { SkipPast("?>"u8.ToArray()); return TagSkipped; }
        if (next == '!')
        {
            if (_reader.PeekAt(2) == '-' && _reader.PeekAt(3) == '-') SkipPast("-->"u8.ToArray());
            else if (_reader.PeekAt(2) == '[') SkipPast("]]>"u8.ToArray());
            else SkipTag();
            return TagSkipped;
        }
        return TagStart;
    }

    private bool IsWrapper()
        => (_sectionOpen is not null && _reader.Matches(_sectionOpen) && IsNameEnd(_reader.PeekAt(_sectionOpen.Length)))
        || (_subSectionOpen is not null && _reader.Matches(_subSectionOpen) && IsNameEnd(_reader.PeekAt(_subSectionOpen.Length)));

    private static bool IsNameEnd(int b) => b is '>' or '/' or ' ' or '\t' or '\r' or '\n';

    /// <summary>Moves past the tag at the position (quotes respected); true when it was self-closing ("/&gt;").</summary>
    private bool SkipTag() => _reader.SkipTag();

    private void SkipPast(byte[] terminator)
    {
        _reader.Skip(1);
        while (true)
        {
            if (!_reader.FindByte(terminator[0])) return;
            if (_reader.Matches(terminator)) { _reader.Skip(terminator.Length); return; }
            _reader.Skip(1);
        }
    }
}

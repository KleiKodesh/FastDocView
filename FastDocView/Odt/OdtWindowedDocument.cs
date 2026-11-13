using System.Xml.Linq;
using FastDocView.Html;

namespace FastDocView.Odt;

/// <summary>
/// An OpenDocument text (.odt / .ott package, or flat .fodt) opened for windowed rendering, behind
/// the same <see cref="IWordWindowedDocument"/> contract as .docx and .doc.
///   • package, in-memory mode (small): one pass inflates content.xml and keeps it for the session;
///   • package, sidecar mode: content.xml checkpoints + the index are saved; reopen reads no body;
///   • flat: the file is the random-access view; the index (sidecar when big) only saves the pass.
/// A sidecar first open with <see cref="WordIndexPolicy.IndexesOnDemand"/> returns after the first
/// <see cref="OnDemandFirstBlockCount"/> blocks and extends the index as blocks are asked for
/// (<see cref="OdtIncrementalIndexBuilder"/>); the finished index is saved when the scan completes.
/// </summary>
public sealed class OdtWindowedDocument : IWordWindowedDocument, IWordDocumentSearchScan
{
    public const int OnDemandFirstBlockCount = 100;

    private OdtRandomAccessIndex? _index;                // null while an on-demand open is still scanning
    private OdtIncrementalIndexBuilder? _builder;
    private IOdtBlockTable _blocks;
    private readonly string? _sidecarPathToSave;
    private readonly OdtBlockXmlReader _reader;
    private readonly OdtStyleCascade _styles;
    private readonly OdtBlockHtmlRenderer _renderer;
    private WordDocumentBlockRenderGuard? _blockRenderGuardField;
    private WordDocumentBlockRenderGuard _blockRenderGuard => _blockRenderGuardField ??= new WordDocumentBlockRenderGuard(FilePath);
    private readonly object _readLock = new();
    private int _footnoteCount, _endnoteCount, _noteReferencesCounted;
    private WordDocumentFileStream? _releasableFile;   // set by Open(path): released after every window
    private bool _isReadingFromFile;                    // false when content.xml was inflated in memory (small package)

    public string FilePath { get; }

    /// <summary>The blocks known so far while an on-demand index is scanning (only grows), exact once complete.</summary>
    public int BlockCount => KnownBlockCount;

    public int ScanAhead(int blocks)
    {
        lock (_readLock)
        {
            if (_builder is not { } builder) return _index!.BlockCount;
            if (_releasableFile is not null && _releasableFile.HasSourceChanged) throw new WordDocumentSourceChangedException(FilePath);
            try
            {
                builder.AdvanceUntilKnown((int)Math.Min(int.MaxValue, (long)builder.KnownBlockCount + Math.Max(1, blocks)));
                if (builder.IsEveryBlockKnown) CompleteIndexLocked();
                else CountNewNoteReferences();
                return _builder?.KnownBlockCount ?? _index!.BlockCount;
            }
            finally { _releasableFile?.Release(); }
        }
    }

    public bool IsIndexComplete => _index is not null;

    public int KnownBlockCount
    {
        get { lock (_readLock) return _builder?.KnownBlockCount ?? _index!.BlockCount; }
    }

    public bool HasSourceChanged => _releasableFile?.HasSourceChanged ?? false;
    public bool IsRightToLeft { get; }
    public string Stylesheet { get; }
    public WordIndexDecision Decision { get; }

    /// <summary>The finished index. Throws while an on-demand open is still scanning (see <see cref="CompleteIndex"/>).</summary>
    public OdtRandomAccessIndex Index => _index ?? throw new InvalidOperationException("The index is still being built on demand; call CompleteIndex() first.");

    /// <summary>Uncompressed body size: content.xml (package) or the whole file (flat).</summary>
    public static long TextBytes(Stream file, bool isFlat) => isFlat ? file.Length : OdtPackage.LocateContentEntry(file).UncompressedSize;

    public static WordIndexDecision Decide(string filePath, WordIndexPolicy policy, bool isFlat)
    {
        using FileStream file = OpenRead(filePath);
        bool hasCachedIndex = OdtRandomAccessIndex.TryLoad(policy.IndexPathFor(filePath), filePath) is not null;
        return policy.Decide(filePath, TextBytes(file, isFlat), hasCachedIndex);
    }

    public static OdtWindowedDocument Open(string filePath, bool isFlat, WordIndexPolicy? policy = null,
                                           IProgress<WordDocumentIndexProgress>? progress = null, CancellationToken cancellation = default)
    {
        policy ??= WordIndexPolicy.Default;
        var info = new FileInfo(filePath);
        string indexPath = policy.IndexPathFor(filePath);
        OdtRandomAccessIndex? cached = OdtRandomAccessIndex.TryLoad(indexPath, filePath);
        var file = new WordDocumentFileStream(filePath);
        try
        {
            WordIndexDecision decision = policy.Decide(filePath, TextBytes(file, isFlat), cached is not null);
            bool isOnDemand = decision.Mode == WordIndexMode.Sidecar && cached is null && policy.IndexesOnDemand;
            OdtWindowedDocument document = isOnDemand
                ? OpenOnDemand(file, filePath, isFlat, decision, info.Length, info.LastWriteTimeUtc.Ticks, progress, indexPath)
                : OpenFromStream(file, filePath, isFlat, decision, cached, info.Length, info.LastWriteTimeUtc.Ticks, progress, cancellation);
            if (!isOnDemand && decision.Mode == WordIndexMode.Sidecar && cached is null)
            {
                try { document.Index.Save(indexPath); }
                catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }   // the sidecar is a cache; the open document stands
            }
            document._releasableFile = file;
            document._isReadingFromFile = isFlat || decision.Mode == WordIndexMode.Sidecar
                                          || (cached is not null && cached.MatchesSource(info.Length, info.LastWriteTimeUtc.Ticks));
            file.Release();                                   // no handle on the file between windows (Word / LibreOffice must be able to save it)
            return document;
        }
        catch
        {
            file.Dispose();
            throw;
        }
    }

    /// <summary>
    /// Host-agnostic entry. <paramref name="file"/> must be seekable; the document takes ownership: it
    /// stays open for the document's lifetime and is disposed with it, except in package in-memory
    /// mode, where it is read and disposed during this call.
    /// </summary>
    public static OdtWindowedDocument OpenFromStream(Stream file, string displayPath, bool isFlat, WordIndexDecision decision,
                                                     OdtRandomAccessIndex? cachedIndex, long sourceLength, long sourceLastWriteTicksUtc,
                                                     IProgress<WordDocumentIndexProgress>? progress = null, CancellationToken cancellation = default)
    {
        if (!isFlat) OdtPackage.ThrowWhenEncrypted(file);
        if (cachedIndex is not null && cachedIndex.MatchesSource(sourceLength, sourceLastWriteTicksUtc) && cachedIndex.IsFlat == isFlat)
            return OverIndex(displayPath, file, cachedIndex, decision);

        // the index pass needs every style (list styles for the labels); the session keeps only what rendering needs
        OdtStyleCascade indexingStyles = IndexingStyles(file, isFlat);

        if (isFlat)
            return OverIndex(displayPath, file, OdtRandomAccessIndexBuilder.BuildFromFlat(file, indexingStyles, sourceLength, sourceLastWriteTicksUtc, progress, cancellation), decision);

        if (decision.Mode == WordIndexMode.InMemory)
        {
            var content = new MemoryStream();
            OdtRandomAccessIndex index = OdtRandomAccessIndexBuilder.BuildFromPackage(file, indexingStyles, sourceLength, sourceLastWriteTicksUtc, progress, long.MaxValue, content, cancellation);
            string? generator = OdtPackage.ReadGenerator(file, isFlat: false);
            file.Dispose();
            return new OdtWindowedDocument(displayPath, index, null, OdtBlockXmlReader.OverBytes(content.ToArray(), index),
                                           RenderingStyles(index.RootOpenTag, index.StyleSections), decision, sidecarPathToSave: null, generator);
        }

        OdtRandomAccessIndex built = OdtRandomAccessIndexBuilder.BuildFromPackage(file, indexingStyles, sourceLength, sourceLastWriteTicksUtc, progress,
                                                                                OdtRandomAccessIndexBuilder.DefaultCheckpointSpanBytes, null, cancellation);
        return OverIndex(displayPath, file, built, decision);
    }

    private static OdtStyleCascade IndexingStyles(Stream file, bool isFlat)
    {
        var styles = new OdtStyleCascade();
        if (isFlat) styles.PageSetup = OdtPackage.ReadPageSetup(file, isFlat: true);
        else OdtPackage.ReadStyles(file, styles);          // one pass: named styles, list styles, page setup
        return styles;
    }

    private static OdtWindowedDocument OverIndex(string displayPath, Stream file, OdtRandomAccessIndex index, WordIndexDecision decision)
        => new(displayPath, index, null, index.IsFlat ? OdtBlockXmlReader.OverFlatFile(file, index) : OdtBlockXmlReader.OverPackage(file, index),
               RenderingStyles(index.RootOpenTag, index.StyleSections), decision, sidecarPathToSave: null, OdtPackage.ReadGenerator(file, index.IsFlat));

    /// <summary>
    /// Returns once the first blocks are known; the rest of the index is scanned as blocks are asked
    /// for. The style sections the renderer needs (font faces, automatic styles) precede office:body,
    /// so they are complete by the time the first block is.
    /// </summary>
    private static OdtWindowedDocument OpenOnDemand(Stream file, string displayPath, bool isFlat, WordIndexDecision decision, long sourceLength,
                                                    long sourceLastWriteTicksUtc, IProgress<WordDocumentIndexProgress>? progress, string sidecarPathToSave)
    {
        if (!isFlat) OdtPackage.ThrowWhenEncrypted(file);
        var builder = new OdtIncrementalIndexBuilder(file, isFlat, IndexingStyles(file, isFlat), sourceLength, sourceLastWriteTicksUtc, progress);
        builder.AdvanceUntilKnown(OnDemandFirstBlockCount);
        if (builder.KnownBlockCount == 0)
        {
            // an empty or malformed body: the one-go rules decide (keep the complete blocks, or throw)
            OdtWindowedDocument whole = OverIndex(displayPath, file, builder.Complete(), decision);
            try { whole.Index.Save(sidecarPathToSave); }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }
            return whole;
        }
        var document = new OdtWindowedDocument(displayPath, null, builder, OdtBlockXmlReader.OverIncrementalBuild(file, builder),
                                               RenderingStyles(builder.StyleSectionElementsForRendering()), decision, sidecarPathToSave,
                                               OdtPackage.ReadGenerator(file, isFlat));
        if (builder.IsEveryBlockKnown) document.CompleteIndexLocked();   // a short body: done already
        return document;
    }

    /// <summary>A first open's rendering cascade: the same filtered sections the index stores, as trees (no string round trip).</summary>
    private static OdtStyleCascade RenderingStyles(XElement[] styleSections)
    {
        var styles = new OdtStyleCascade();
        foreach (XElement section in styleSections) styles.AddSection(section);
        return styles;
    }

    /// <summary>The cascade from captured style sections (no list styles) — a reopen never touches styles.xml.</summary>
    private static OdtStyleCascade RenderingStyles(string rootOpenTag, string[] styleSections)
    {
        var styles = new OdtStyleCascade();
        foreach (string section in styleSections) styles.AddSection(OdtBodyBlockScanner.ParseSection(rootOpenTag, section));
        return styles;
    }

    private OdtWindowedDocument(string filePath, OdtRandomAccessIndex? index, OdtIncrementalIndexBuilder? builder, OdtBlockXmlReader reader,
                                OdtStyleCascade styles, WordIndexDecision decision, string? sidecarPathToSave, string? generator)
    {
        FilePath = filePath;
        _index = index;
        _builder = builder;
        _blocks = index is not null ? reader.Table : builder!;
        _sidecarPathToSave = sidecarPathToSave;
        _reader = reader;
        _styles = styles;
        _renderer = new OdtBlockHtmlRenderer(styles) { Generator = generator };
        Decision = decision;
        CountNewNoteReferences();
        IsRightToLeft = styles.IsDocumentRightToLeft;
        Stylesheet = WordDocumentHtmlStylesheet.Build(styles.DefaultParagraph.Formatting.RunFormatting, IsRightToLeft);
        if (index is not null) RecordConnectedShapes(index);
    }

    /// <summary>
    /// Index time, once the blocks are final: the geometry of every shape with an id (snap rectangle, glue points), read from the few blocks
    /// that hold one (block by block, windowed), so a connector glued to a shape in another block routes without reading that block.
    /// </summary>
    private void RecordConnectedShapes(OdtRandomAccessIndex index)
    {
        if (index.ConnectedShapes is null)
        {
            var shapes = new Dictionary<string, OdtConnectedShape>(StringComparer.Ordinal);
            _renderer.RootOpenTag ??= _blocks.RootOpenTag;
            _renderer.Pictures ??= ReadPicture;
            try
            {
                foreach (int blockIndex in index.ConnectedShapeBlocks)
                    foreach ((int read, XElement element) in _reader.ReadBlocks(blockIndex, 1))
                        _renderer.RecordConnectedShapes(element, read, shapes);
            }
            finally { _releasableFile?.Release(); }
            index.ConnectedShapes = new OdtConnectedShapeTable(shapes);
        }
        _renderer.ConnectedShapes = index.ConnectedShapes.Find;
    }

    /// <inheritdoc/>
    /// <remarks>
    /// "p:" ids: the package entry, read from the file by path on a short shared open (never blocks Word / LibreOffice) -
    /// only an entry with a picture extension a browser shows, so no other part of the package can be asked for. "b:" ids
    /// (flat files): the one block is read again and its n-th office:binary-data picture decoded.
    /// </remarks>
    public WordDocumentPicture? ReadPicture(string pictureId)
    {
        if (pictureId is null) return null;
        if (pictureId.StartsWith(OdtBlockHtmlRenderer.PackagePicturePrefix, StringComparison.Ordinal))
        {
            string partName = pictureId.Substring(OdtBlockHtmlRenderer.PackagePicturePrefix.Length);
            if (WordPictureHtml.ContentTypeOf(partName) is null || !File.Exists(FilePath)) return null;
            lock (_readLock)
            {
                using FileStream file = WordDocumentFileStream.OpenShared(FilePath);
                if (OdtPackage.IsFlatOpenDocumentText(file)) return null;
                file.Position = 0;
                using var zip = new System.IO.Compression.ZipArchive(file, System.IO.Compression.ZipArchiveMode.Read, leaveOpen: true);
                System.IO.Compression.ZipArchiveEntry? entry = zip.GetEntry(partName);
                if (entry is null) return null;
                using Stream data = entry.Open();
                var bytes = new MemoryStream();
                data.CopyTo(bytes);
                return WordPictureHtml.Picture(partName, bytes.ToArray());
            }
        }
        if (pictureId.StartsWith(OdtBlockHtmlRenderer.FillImagePrefix, StringComparison.Ordinal))
        {
            // a fill image stored in the styles (draw:fill-image office:binary-data): the kept style holds it
            XElement? image = _styles.DrawingStyle("fill-image", pictureId.Substring(OdtBlockHtmlRenderer.FillImagePrefix.Length));
            if (image?.Element(OdtStyleCascade.Office + "binary-data") is not { } data || OdtBlockHtmlRenderer.BinaryPictureBytes(data) is not { } fill) return null;
            return WordPictureHtml.ExtensionOf(fill).Length > 0 ? WordPictureHtml.Picture(WordPictureHtml.ExtensionOf(fill), fill) : null;
        }
        if (!pictureId.StartsWith(OdtBlockHtmlRenderer.BinaryPicturePrefix, StringComparison.Ordinal)) return null;
        string[] parts = pictureId.Substring(OdtBlockHtmlRenderer.BinaryPicturePrefix.Length).Split(':');
        if (parts.Length != 2 || !int.TryParse(parts[0], System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int blockIndex)
            || !int.TryParse(parts[1], System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int ordinal))
            return null;
        lock (_readLock)
        {
            if (blockIndex >= _blocks.BlockCount) return null;
            if (_releasableFile is not null && _isReadingFromFile && _releasableFile.HasSourceChanged) throw new WordDocumentSourceChangedException(FilePath);
            try
            {
                foreach ((_, XElement element) in _reader.ReadBlocks(blockIndex, 1))
                {
                    XElement? image = OdtBlockHtmlRenderer.BinaryPictureImages(element).Skip(ordinal).FirstOrDefault();
                    if (image?.Element(OdtStyleCascade.Office + "binary-data") is not { } binary || OdtBlockHtmlRenderer.BinaryPictureBytes(binary) is not { } bytes) return null;
                    return WordPictureHtml.Picture(WordPictureHtml.ExtensionOf(bytes), bytes);
                }
                return null;
            }
            finally { _releasableFile?.Release(); }
        }
    }

    private static readonly byte[] ObjectElementBytes = System.Text.Encoding.ASCII.GetBytes(":object");

    /// <summary>Where a "prefix:object" start tag begins in raw block bytes (the name followed by white space, "/" or ">"), or -1.</summary>
    private static int ObjectElementAt(ReadOnlySpan<byte> bytes)
    {
        for (int at = 0; at < bytes.Length;)
        {
            int found = bytes.Slice(at).IndexOf(ObjectElementBytes);
            if (found < 0) return -1;
            int end = at + found + ObjectElementBytes.Length;
            if (end < bytes.Length && bytes[end] is 0x20 or 0x09 or 0x0D or 0x0A or 0x2F or 0x3E) return at + found;
            at += found + 1;
        }
        return -1;
    }

    /// <summary>The block whose bytes hold an offset of the part (the last block of [first, end) starting at or before it).</summary>
    private static int BlockAtOffset(IOdtBlockTable blocks, long offset, int first, int end)
    {
        int low = first, high = end - 1;
        while (low < high)
        {
            int middle = low + (high - low + 1) / 2;
            if (blocks.BlockStartOffset(middle) <= offset) low = middle; else high = middle - 1;
        }
        return low;
    }

    /// <summary>A block's chart texts fed to the matcher, its hits added to the batch's (merged with the block's own text hits).</summary>
    private void FeedChartTexts(int blockIndex, WordDocumentSearchMatcher matcher, List<int> hits, List<int> counts)
    {
        int at = hits.BinarySearch(blockIndex);
        matcher.Reset();
        if (matcher.IsCapturingText && at >= 0) matcher.ResumeCapture(counts[at]);   // the charts' text joins the block's own
        foreach ((_, XElement element) in _reader.ReadBlocks(blockIndex, 1))
            foreach (string text in ChartTexts(element))
            {
                matcher.Separator();
                matcher.Feed(text.AsSpan());
            }
        int count = matcher.Count;
        matcher.Reset();
        if (count == 0) return;
        if (at >= 0) { if (!matcher.IsCapturingText) counts[at] += count; }
        else { hits.Insert(~at, blockIndex); counts.Insert(~at, count); }
    }

    // ── charts: an embedded chart sub-document read when its block is rendered or searched ──

    private const int ChartCacheSize = 16;
    private readonly Dictionary<string, DocxHtml.DocxChart?> _charts = new(StringComparer.Ordinal);

    /// <summary>
    /// The chart of a draw:object (<see cref="OdtChart"/>): a flat file's inline office:document, else the package's
    /// "Object N/content.xml" the xlink:href names, read on a short shared open only when the block is rendered or searched
    /// (a few recent ones kept). Null when the object is not a chart (a formula, a spreadsheet: their preview picture shows).
    /// </summary>
    internal DocxHtml.DocxChart? ReadChart(XElement embedded)
    {
        if (embedded.Element(OdtStyleCascade.Office + "document") is { } inline) return OdtChart.Read(inline);
        if ((string?)embedded.Attribute(OdtBlockHtmlRenderer.XlinkNs + "href") is not { Length: > 0 } href || href.Contains(':')) return null;
        string folder = (href.StartsWith("./", StringComparison.Ordinal) ? href.Substring(2) : href).TrimEnd('/');
        if (folder.Length == 0 || folder.Contains("..")) return null;
        lock (_readLock)
        {
            // a replaced file: its charts are stale, so the cache goes and the host reopens (as for every other read)
            if (_releasableFile is not null && _releasableFile.HasSourceChanged)
            {
                _charts.Clear();
                throw new WordDocumentSourceChangedException(FilePath);
            }
            if (_charts.TryGetValue(folder, out DocxHtml.DocxChart? known)) return known;
            DocxHtml.DocxChart? chart = null;
            // a damaged sub-document (bad zip entry, malformed XML) shows the preview picture instead of failing the block or the search
            try
            {
            if (File.Exists(FilePath))
            {
                using FileStream file = WordDocumentFileStream.OpenShared(FilePath);
                if (!OdtPackage.IsFlatOpenDocumentText(file))
                {
                    file.Position = 0;
                    using var zip = new System.IO.Compression.ZipArchive(file, System.IO.Compression.ZipArchiveMode.Read, leaveOpen: true);
                    if (zip.GetEntry(folder + "/content.xml") is { } entry && IsChartContent(entry))
                    {
                        using Stream data = entry.Open();
                        var settings = new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Prohibit, XmlResolver = null };
                        using var xml = System.Xml.XmlReader.Create(data, settings);
                        if (XDocument.Load(xml).Root is { } root) chart = OdtChart.Read(root);
                    }
                }
            }
            }
            catch (Exception exception) when (exception is InvalidDataException or System.Xml.XmlException or FormatException or OverflowException) { chart = null; }
            if (_charts.Count >= ChartCacheSize) _charts.Clear();
            _charts[folder] = chart;
            return chart;
        }
    }

    /// <summary>True when a sub-document's office:body holds office:chart (read up to the body's first child, so a spreadsheet or formula is never loaded).</summary>
    private static bool IsChartContent(System.IO.Compression.ZipArchiveEntry entry)
    {
        using Stream data = entry.Open();
        var settings = new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Prohibit, XmlResolver = null };
        using var xml = System.Xml.XmlReader.Create(data, settings);
        string office = OdtStyleCascade.Office.NamespaceName;
        bool inBody = false;
        while (xml.Read())
        {
            if (xml.NodeType != System.Xml.XmlNodeType.Element) continue;
            if (!inBody) { inBody = xml.LocalName == "body" && xml.NamespaceURI == office; continue; }
            return xml.LocalName == "chart" && xml.NamespaceURI == office;
        }
        return false;
    }

    /// <summary>The texts of a block's charts (title, axis titles, categories, legend, data labels) and formulas, for search; empty when it has none.</summary>
    private IEnumerable<string> ChartTexts(XElement block)
    {
        foreach (XElement embedded in block.Descendants(OdtBlockHtmlRenderer.DrawNs + "object"))
        {
            if (ReadFormula(embedded) is { } formula) { yield return WordMathMlHtml.Text(formula); continue; }
            if (embedded.Element(OdtStyleCascade.Office + "document") is null && ReadChart(embedded) is { } chart)
                foreach (string text in chart.Texts()) yield return text;
        }
    }

    // ── formulas: an embedded LibreOffice Math object's MathML, read when its block is rendered or searched ──

    private readonly Dictionary<string, XElement?> _formulas = new(StringComparer.Ordinal);

    /// <summary>
    /// The MathML math:math of a formula object: a flat file's inline math:math inside the draw:object, else the package's
    /// "Object N/content.xml" (its root is math:math - LibreOffice's Math export, starmath/source/mathml/mathmlexport.cxx),
    /// read on a short shared open only when the block is rendered or searched (a few recent ones kept). Null: not a formula.
    /// </summary>
    internal XElement? ReadFormula(XElement embedded)
    {
        if (embedded.Elements().FirstOrDefault(e => e.Name.LocalName == "math" && e.Name.NamespaceName == WordMathMlHtml.MathMlNamespace) is { } inline) return inline;
        if (embedded.Element(OdtStyleCascade.Office + "document") is not null) return null;
        if ((string?)embedded.Attribute(OdtBlockHtmlRenderer.XlinkNs + "href") is not { Length: > 0 } href || href.Contains(':')) return null;
        string folder = (href.StartsWith("./", StringComparison.Ordinal) ? href.Substring(2) : href).TrimEnd('/');
        if (folder.Length == 0 || folder.Contains("..")) return null;
        lock (_readLock)
        {
            if (_formulas.TryGetValue(folder, out XElement? known)) return known;
            XElement? formula = null;
            try
            {
                if (File.Exists(FilePath))
                {
                    using FileStream file = WordDocumentFileStream.OpenShared(FilePath);
                    if (!OdtPackage.IsFlatOpenDocumentText(file))
                    {
                        file.Position = 0;
                        using var zip = new System.IO.Compression.ZipArchive(file, System.IO.Compression.ZipArchiveMode.Read, leaveOpen: true);
                        if (zip.GetEntry(folder + "/content.xml") is { } entry && IsFormulaContent(entry))
                        {
                            using Stream data = entry.Open();
                            var settings = new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Ignore, XmlResolver = null };
                            using var xml = System.Xml.XmlReader.Create(data, settings);
                            formula = XDocument.Load(xml).Root;
                        }
                    }
                }
            }
            catch (Exception exception) when (exception is InvalidDataException or System.Xml.XmlException or FormatException or OverflowException) { formula = null; }
            if (_formulas.Count >= ChartCacheSize) _formulas.Clear();
            _formulas[folder] = formula;
            return formula;
        }
    }

    /// <summary>True when a sub-document's root element is MathML math:math (read up to the root only).</summary>
    private static bool IsFormulaContent(System.IO.Compression.ZipArchiveEntry entry)
    {
        using Stream data = entry.Open();
        var settings = new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Ignore, XmlResolver = null };
        using var xml = System.Xml.XmlReader.Create(data, settings);
        while (xml.Read())
            if (xml.NodeType == System.Xml.XmlNodeType.Element) return xml.LocalName == "math" && xml.NamespaceURI == WordMathMlHtml.MathMlNamespace;
        return false;
    }

    private IReadOnlyList<OdtNoteReference> KnownNoteReferences => (IReadOnlyList<OdtNoteReference>?)_index?.NoteReferences ?? _builder!.KnownNoteReferences;

    /// <summary>Adds the note references found since the last call to the per-kind counts.</summary>
    private void CountNewNoteReferences()
    {
        IReadOnlyList<OdtNoteReference> references = KnownNoteReferences;
        for (; _noteReferencesCounted < references.Count; _noteReferencesCounted++)
            if (references[_noteReferencesCounted].Kind == WordDocumentNoteKind.Footnote) _footnoteCount++; else _endnoteCount++;
    }

    /// <summary>
    /// Finishes an on-demand index now (scans the rest of the body) and saves the sidecar when the
    /// document was opened from a path. No-op when already complete.
    /// </summary>
    public OdtRandomAccessIndex CompleteIndex()
    {
        lock (_readLock) return CompleteIndexLocked();
    }

    private OdtRandomAccessIndex CompleteIndexLocked()
    {
        if (_index is not null) return _index;
        OdtRandomAccessIndex index = _builder!.Complete();
        // Swap the scan's live tables for the finished index (same offsets, same checkpoints) and free the scanner.
        _index = index;
        _blocks = _reader.SwitchToIndex(index);
        _builder = null;
        RecordConnectedShapes(index);
        _noteReferences = null;
        CountNewNoteReferences();
        if (_sidecarPathToSave is not null)
        {
            try { index.Save(_sidecarPathToSave); }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { /* reading goes on; the next open rebuilds */ }
        }
        return index;
    }

    /// <summary>Past the end (a count read before the index completed) it answers 0 instead of throwing.</summary>
    public int EstimatedTextLength(int blockIndex)
    {
        lock (_readLock)
        {
            if (_builder is { } builder) return blockIndex < builder.KnownBlockCount ? builder.BlockTextLengthOf(blockIndex) : builder.AverageTextLength;
            // exact once the block's span has been rescanned (a window near it was read), else its 64-block group's average
            return blockIndex < _index!.BlockCount ? ((OdtCompactBlockTable)_blocks).KnownTextLength(blockIndex) ?? _index.EstimatedTextLength(blockIndex) : 0;
        }
    }

    /// <summary>The block's list label as the index pass counted it (null when not a numbered list paragraph) - rescanned from its span.</summary>
    public string? ListLabelOf(int blockIndex) { lock (_readLock) return WithFileReleased(() => _blocks.ListLabelOf(blockIndex)); }
    /// <summary>The list depth (0-based) of a list paragraph, or -1.</summary>
    public int ListLevelOf(int blockIndex) { lock (_readLock) return WithFileReleased(() => _blocks.ListLevelOf(blockIndex)); }
    public WordDocumentBlockKind BlockKindOf(int blockIndex) { lock (_readLock) return WithFileReleased(() => _blocks.BlockKindOf(blockIndex)); }

    private T WithFileReleased<T>(Func<T> lookup)
    {
        if (_releasableFile is not null && _isReadingFromFile && _releasableFile.HasSourceChanged) throw new WordDocumentSourceChangedException(FilePath);
        try { return lookup(); }
        finally { _releasableFile?.Release(); }
    }

    public IReadOnlyList<WordDocumentBlockHtml> RenderBlocks(int firstBlockIndex, int count)
    {
        lock (_readLock)
        {
            if (_releasableFile is not null && _isReadingFromFile && _releasableFile.HasSourceChanged)
                throw new WordDocumentSourceChangedException(FilePath);
            try
            {
                if (_builder is { } builder)
                {
                    builder.AdvanceUntilKnown(firstBlockIndex + count);     // the on-demand scan reopens the released file like any read
                    if (builder.IsEveryBlockKnown) CompleteIndexLocked();
                    else CountNewNoteReferences();
                }
                return _blockRenderGuard.Render(firstBlockIndex, count, _blocks.BlockCount, RenderBlocksLocked, _blocks.BlockKindOf);
            }
            finally { _releasableFile?.Release(); }
        }
    }

    private List<WordDocumentBlockHtml> RenderBlocksLocked(int firstBlockIndex, int count)
    {
        var result = new List<WordDocumentBlockHtml>(Math.Max(0, count));
        foreach ((int blockIndex, XElement element) in _reader.ReadBlocks(firstBlockIndex, count))
        {
            _renderer.CurrentBlockIndex = blockIndex;
            _renderer.ColumnWidthPt = WordDocumentPageSetup.TextWidthPointsAt(Sections, blockIndex);
            _renderer.ColumnLeftPt = WordDocumentPageSetup.ColumnLeftPointsAt(Sections, blockIndex);
            _renderer.RootOpenTag ??= _blocks.RootOpenTag;
            _renderer.Pictures ??= ReadPicture;
            _renderer.Charts ??= ReadChart;
            _renderer.Formulas ??= ReadFormula;
            _renderer.ListStyles ??= _blocks.ListStyleNamed;
            _renderer.ResetProvisional();
            (string html, IReadOnlyList<WordDocumentNoteHtml> notes) = _renderer.Render(element, _blocks.ListLabelOf(blockIndex), _blocks.ListLevelOf(blockIndex), _blocks.ListLabelLayoutOf(blockIndex),
                                                                                       _blocks.NoteReferencesInBlocks(blockIndex, blockIndex), null);
            result.Add(new WordDocumentBlockHtml(blockIndex, _blocks.BlockKindOf(blockIndex), html, notes, _renderer.LastBlockIsProvisional));
        }
        return result;
    }

    /// <summary>
    /// Scans blocks [fromBlock, fromBlock + maxBlocksToScan): one contiguous read of their bytes, each block's paragraph
    /// text fed to the matcher straight from the bytes (no XElement, no rendering). Notes are inline in ODF
    /// (text:note-body) and are skipped: the block's HTML shows only the note's mark. An on-demand
    /// index is extended as far as asked, as <see cref="RenderBlocks"/> does.
    /// </summary>
    public WordDocumentSearchBatch Search(string query, WordDocumentSearchOptions options, int fromBlock, int maxBlocksToScan)
        => ((IWordDocumentSearchScan)this).ScanBlocks(new WordDocumentSearchMatcher(query, options), fromBlock, maxBlocksToScan);

    /// <summary>The product block scan, shared with <see cref="WordDocumentText"/> (one text path, <see cref="IWordDocumentSearchScan"/>).</summary>
    WordDocumentSearchBatch IWordDocumentSearchScan.ScanBlocks(WordDocumentSearchMatcher matcher, int fromBlock, int maxBlocksToScan)
        => SearchWith(matcher, fromBlock, maxBlocksToScan, OdtSearchTextMode.Bytes);

    /// <summary>How a search batch reads block text: the byte scanner (the product) or the XmlReader reference over the index ranges (golden checks).</summary>
    internal enum OdtSearchTextMode { Bytes, ReferenceXmlReader }

    /// <summary>Blocks per raw search read: bounds the bytes one read holds whatever the batch size.</summary>
    private const int SearchPieceBlocks = 1000;

    internal WordDocumentSearchBatch Search(string query, WordDocumentSearchOptions options, int fromBlock, int maxBlocksToScan, OdtSearchTextMode mode)
        => SearchWith(new WordDocumentSearchMatcher(query, options), fromBlock, maxBlocksToScan, mode);

    /// <summary>The same scan with a caller's matcher (<see cref="WordDocumentText"/> captures the text through it).</summary>
    internal WordDocumentSearchBatch SearchWith(WordDocumentSearchMatcher matcher, int fromBlock, int maxBlocksToScan, OdtSearchTextMode mode)
    {
        if (matcher.IsEmpty) return WordDocumentSearchBatch.Nothing;
        fromBlock = Math.Max(0, fromBlock);
        lock (_readLock)
        {
            if (_releasableFile is not null && _isReadingFromFile && _releasableFile.HasSourceChanged)
                throw new WordDocumentSourceChangedException(FilePath);
            try
            {
                int wantedEnd = (int)Math.Min(int.MaxValue, (long)fromBlock + Math.Max(1, maxBlocksToScan));
                if (_builder is { } builder)
                {
                    builder.AdvanceUntilKnown(wantedEnd);
                    if (builder.IsEveryBlockKnown) CompleteIndexLocked();
                    else CountNewNoteReferences();
                }
                int end = Math.Min(_blocks.BlockCount, wantedEnd);
                var hits = new List<int>();
                var counts = new List<int>();
                var objectBlocks = new SortedSet<int>();   // blocks holding a draw:object: their charts' texts are searched after the scan
                if (fromBlock < end)
                {
                    _searchText ??= new OdtSearchTextScanner(_blocks.RootOpenTag);
                    long pieceStartOffset = -1;            // a byte piece's start in the part while it is scanned: hit block starts are recorded
                    var pieceHitStarts = new List<long>();  // to verify the scanner's numbering if the piece fails
                    void BlockEnded(int blockIndex)
                    {
                        if (matcher.Count > 0)
                        {
                            hits.Add(blockIndex);
                            counts.Add(matcher.Count);
                            if (pieceStartOffset >= 0) pieceHitStarts.Add(pieceStartOffset + _searchText!.CurrentBlockStart);
                        }
                        matcher.Reset();
                    }
                    void FeedRanges(int first, int stop)
                    {
                        (byte[] bytes, List<(int blockIndex, int start, int length)> blocks) = _reader.ReadBlockBytes(first, stop - first);
                        if (mode == OdtSearchTextMode.ReferenceXmlReader) _referenceSearchText ??= new WordXmlSearchText(_blocks.RootOpenTag);
                        foreach ((int blockIndex, int start, int length) in blocks)
                        {
                            matcher.Reset();
                            if (mode == OdtSearchTextMode.ReferenceXmlReader) _referenceSearchText!.FeedOdf(bytes, start, length, matcher);
                            else _searchText.Feed(bytes, start, length, matcher);
                            BlockEnded(blockIndex);
                            if (ObjectElementAt(bytes.AsSpan(start, length)) >= 0) objectBlocks.Add(blockIndex);
                        }
                    }
                    int scanFrom = fromBlock;
                    // the finished index: the scanner finds the block boundaries in raw reads (no span scan), a bounded
                    // piece at a time so a huge batch never holds more than a piece's bytes
                    if (_blocks is OdtCompactBlockTable compact && mode == OdtSearchTextMode.Bytes)
                    {
                        while (scanFrom < end)
                        {
                            if (compact.SearchStart(scanFrom, out long startOffset) is not { } frames)
                            {
                                // not where a byte run can start (a batch begun out of order): per-block ranges up to the next anchor
                                int anchorBlock = Math.Min(end, compact.NextSearchStart(scanFrom));
                                FeedRanges(scanFrom, anchorBlock);
                                scanFrom = anchorBlock;
                                continue;
                            }
                            int pieceEnd = Math.Min(end, scanFrom + SearchPieceBlocks);
                            (byte[] buffer, int start, int length) = compact.ReadSearchBytes(startOffset, pieceEnd, out bool isEndExact);
                            int hitsBefore = hits.Count;
                            pieceHitStarts.Clear();
                            pieceStartOffset = startOffset;
                            matcher.Reset();
                            WordSearchBlocksResult result = _searchText.FeedBlocks(buffer.AsSpan(start, length), scanFrom, pieceEnd, frames, matcher, BlockEnded, out long stoppedAt);
                            pieceStartOffset = -1;
                            // chart objects in the piece's own blocks (bytes past pieceEnd's start belong to later batches); collected
                            // before the verdict, so a fallen-back or redone piece keeps them (the set has each block once)
                            long pieceBytes = pieceEnd < _blocks.BlockCount ? Math.Min(length, compact.BlockStartOffset(pieceEnd) - startOffset) : length;
                            for (int at = 0, found; at < pieceBytes && (found = ObjectElementAt(buffer.AsSpan(start + at, (int)pieceBytes - at))) >= 0; at += found + 1)
                                objectBlocks.Add(BlockAtOffset(compact, startOffset + at + found, scanFrom, pieceEnd));
                            WordSearchPieceVerdict verdict = WordSearchPiece.Judge(result, isEndExact, stoppedAt, length);
                            if (verdict == WordSearchPieceVerdict.FallBackFromFailedBlock)
                            {
                                // the rest block by block, from the index's ranges; hits the scan reported from that block on are dropped.
                                // The hits before it stand only when the scanner's numbering agrees with the index there (the failing
                                // block and every hit block start where the index says); a drift redoes the whole piece.
                                scanFrom = WordSearchPiece.FallBackBlock(scanFrom, (int)stoppedAt, startOffset + _searchText.CurrentBlockStart,
                                                                         hits, pieceHitStarts, hitsBefore, compact.BlockStartOffset);
                                int keep = hits.FindIndex(hitsBefore, hit => hit >= scanFrom);
                                if (keep >= 0) { hits.RemoveRange(keep, hits.Count - keep); counts.RemoveRange(keep, counts.Count - keep); }
                                break;
                            }
                            if (verdict == WordSearchPieceVerdict.RedoPiece)
                            {
                                // the scanner's blocks are not the index's: redo the piece the slow way, so no hit is mis-numbered
                                hits.RemoveRange(hitsBefore, hits.Count - hitsBefore);
                                counts.RemoveRange(hitsBefore, counts.Count - hitsBefore);
                                break;
                            }
                            if (result == WordSearchBlocksResult.Stopped) compact.RememberSearchResume(pieceEnd, startOffset + stoppedAt, _searchText.StoppedFrames());
                            BlocksSearchedFromRawBytes += pieceEnd - scanFrom;
                            scanFrom = pieceEnd;
                        }
                    }
                    if (scanFrom < end) FeedRanges(scanFrom, end);
                    foreach (int blockIndex in objectBlocks) FeedChartTexts(blockIndex, matcher, hits, counts);
                }
                return WordDocumentSearchBatch.Finish(hits, counts, Math.Max(end, fromBlock), _index is not null, _blocks.BlockCount);
            }
            finally
            {
                (_blocks as OdtCompactBlockTable)?.ShrinkSearchBuffer();   // a batch that needed an oversize buffer drops back to the normal size
                _releasableFile?.Release();
            }
        }
    }
    private OdtSearchTextScanner? _searchText;

    /// <summary>The notes search, apart from the blocks' (see <see cref="IWordWindowedDocument.SearchNotes"/>).</summary>
    public WordDocumentSearchBatch SearchNotes(string query, WordDocumentSearchOptions options, int fromNote, int maxNotesToScan)
        => WordDocumentNoteSearch.Search(this, query, options, fromNote, maxNotesToScan);
    /// <summary>Blocks the byte scanner read from raw runs (not per-block ranges) - the tests check the fast path ran.</summary>
    internal long BlocksSearchedFromRawBytes { get; private set; }
    private WordXmlSearchText? _referenceSearchText;

    /// <summary>Headings of the known blocks while an on-demand index is scanning; all of them once complete.</summary>
    public IReadOnlyList<WordDocumentHeading> Headings
    {
        get { lock (_readLock) return _builder is { } builder ? builder.KnownHeadings.ToArray() : _index!.Headings; }
    }

    /// <summary>Page text width per section (master pages, see <see cref="OdtPageSetup"/>); the known ones while scanning.</summary>
    public int? FindBookmark(string name)
    {
        lock (_readLock)
        {
            if (_index is { } index) return index.Bookmarks.BlockOf(name);
            OdtIncrementalIndexBuilder builder = _builder!;
            try
            {
                // scan on only until the bookmark turns up (a target near the start costs nothing more)
                while (true)
                {
                    if (builder.FindKnownBookmark(name) is { } block) return block;
                    if (!builder.AdvanceOne()) break;
                }
                return CompleteIndexLocked().Bookmarks.BlockOf(name);
            }
            finally { _releasableFile?.Release(); }
        }
    }

    public IReadOnlyList<WordDocumentSection> Sections
    {
        get
        {
            lock (_readLock)
            {
                if (_builder is { } builder) return WordDocumentPageSetup.FromSectionStarts(builder.KnownSectionStarts);
                OdtRandomAccessIndex index = _index!;
                return _sections ??= WordDocumentPageSetup.FromSectionStarts(
                    index.SectionFirstBlocks.Select((first, i) => new WordDocumentSection(first, index.SectionTextWidths[i], index.SectionColumnLefts[i] - 1)));
            }
        }
    }
    private IReadOnlyList<WordDocumentSection>? _sections;

    public (int footnotes, int endnotes) NoteCounts
    {
        get { lock (_readLock) return (_footnoteCount, _endnoteCount); }
    }

    public IReadOnlyList<WordDocumentNoteReference> NoteReferences
    {
        get
        {
            lock (_readLock)
            {
                if (_builder is not null)
                    return KnownNoteReferences.Select(r => new WordDocumentNoteReference(r.Kind, r.Number, r.BlockIndex)).ToArray();
                return _noteReferences ??= _index!.NoteReferences.Select(r => new WordDocumentNoteReference(r.Kind, r.Number, r.BlockIndex)).ToArray();
            }
        }
    }
    private WordDocumentNoteReference[]? _noteReferences;

    /// <summary>
    /// Notes are inline in ODF (text:note-body): only the blocks holding the window's notes are rendered. The window lies in
    /// the known references, so the on-demand scan is not extended.
    /// </summary>
    public IReadOnlyList<WordDocumentNoteHtml> RenderNotes(int firstNoteIndex, int count)
    {
        lock (_readLock)
        {
            // all of the window's runs under one lock, one source check and one file open (not one per run)
            IReadOnlyList<WordDocumentNoteReference> references = NoteReferences;
            if (_releasableFile is not null && _isReadingFromFile && _releasableFile.HasSourceChanged)
                throw new WordDocumentSourceChangedException(FilePath);
            try { return WordDocumentNoteWindow.RenderFromBlocks(references, firstNoteIndex, count, RenderBlocksLocked); }
            finally { _releasableFile?.Release(); }
        }
    }

    private static FileStream OpenRead(string path) => WordDocumentFileStream.OpenShared(path);     // Word / LibreOffice may hold the file open

    public void Dispose() => _reader.Dispose();
}

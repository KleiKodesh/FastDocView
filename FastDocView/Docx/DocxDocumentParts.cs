using FastDocView.Html;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using FastDocView.DocxHtml;

namespace FastDocView.Docx;

/// <summary>
/// The small parts of a .docx that are loaded whole (styles, theme, settings, numbering) — a few hundred KB at
/// most, unlike the body. Read through ZipArchive over any seekable stream, so it works the same
/// on a file-system host and in the browser.
/// </summary>
public sealed class DocxDocumentParts
{
    public DocxStyleCascade Styles { get; }
    /// <summary>The main part's hyperlink relationships: r:id → target (external URLs as written). Small; read at open.</summary>
    public IReadOnlyDictionary<string, string> HyperlinkTargets { get; }
    /// <summary>footnotes.xml's and endnotes.xml's own hyperlink relationships (a note's r:id points there, not into document.xml.rels).</summary>
    public IReadOnlyDictionary<string, string> FootnoteHyperlinkTargets { get; }
    public IReadOnlyDictionary<string, string> EndnoteHyperlinkTargets { get; }

    private static Dictionary<string, string> LoadHyperlinkTargets(ZipArchive zip, string relationshipsPartName)
    {
        var targets = new Dictionary<string, string>(StringComparer.Ordinal);
        XElement? root = LoadPart(zip, relationshipsPartName)?.Root;
        if (root is null) return targets;
        foreach (XElement relationship in root.Elements())
        {
            string? type = (string?)relationship.Attribute("Type"), id = (string?)relationship.Attribute("Id"), target = (string?)relationship.Attribute("Target");
            if (id is not null && target is not null && type is not null && type.EndsWith("/hyperlink", StringComparison.OrdinalIgnoreCase)) targets[id] = target;
        }
        return targets;
    }

    /// <summary>The picture relationships of the main part and the note parts: r:id -> package part name (resolved). Small; read at open.</summary>
    public IReadOnlyDictionary<string, string> ImageParts { get; }
    public IReadOnlyDictionary<string, string> FootnoteImageParts { get; }
    public IReadOnlyDictionary<string, string> EndnoteImageParts { get; }

    private static Dictionary<string, string> LoadImageParts(ZipArchive zip, string? sourcePartName)
    {
        var parts = new Dictionary<string, string>(StringComparer.Ordinal);
        XElement? root = sourcePartName is null ? null : LoadPart(zip, DocxPackagePartNames.RelationshipsPartName(sourcePartName))?.Root;
        if (root is null) return parts;
        foreach (XElement relationship in root.Elements())
        {
            string? type = (string?)relationship.Attribute("Type"), id = (string?)relationship.Attribute("Id"), target = (string?)relationship.Attribute("Target");
            if (id is null || string.IsNullOrEmpty(target) || type is null || (string?)relationship.Attribute("TargetMode") == "External") continue;
            if (type.EndsWith("/image", StringComparison.OrdinalIgnoreCase)) parts[id] = DocxPackagePartNames.ResolveTarget(sourcePartName!, target!);
        }
        return parts;
    }

    /// <summary>Every internal relationship of the main part and the note parts: r:id -> (type, package part name). Small; read at open.
    /// The drawing parts a drawing names (chart, diagram data / drawing) are read from here on demand, when the block is rendered.</summary>
    public IReadOnlyDictionary<string, (string Type, string PartName)> RelatedParts { get; }
    public IReadOnlyDictionary<string, (string Type, string PartName)> FootnoteRelatedParts { get; }
    public IReadOnlyDictionary<string, (string Type, string PartName)> EndnoteRelatedParts { get; }

    /// <summary>The internal relationships of any part (its .rels), read now: r:id -> (type, resolved part name).</summary>
    public Dictionary<string, (string Type, string PartName)> ReadRelationshipsOf(string? sourcePartName)
    {
        using var zip = new ZipArchive(_package, ZipArchiveMode.Read, leaveOpen: true);
        return LoadRelatedParts(zip, sourcePartName);
    }

    private static Dictionary<string, (string Type, string PartName)> LoadRelatedParts(ZipArchive zip, string? sourcePartName)
    {
        var parts = new Dictionary<string, (string, string)>(StringComparer.Ordinal);
        XElement? root = sourcePartName is null ? null : LoadPart(zip, DocxPackagePartNames.RelationshipsPartName(sourcePartName))?.Root;
        if (root is null) return parts;
        foreach (XElement relationship in root.Elements())
        {
            string? type = (string?)relationship.Attribute("Type"), id = (string?)relationship.Attribute("Id"), target = (string?)relationship.Attribute("Target");
            if (id is null || string.IsNullOrEmpty(target) || type is null || (string?)relationship.Attribute("TargetMode") == "External") continue;
            parts[id] = (type, DocxPackagePartNames.ResolveTarget(sourcePartName!, target!));
        }
        return parts;
    }

    /// <summary>A package part parsed (a chart or diagram part a drawing names), read on demand; null when absent.</summary>
    public XDocument? ReadPartXml(string partName)
    {
        using var zip = new ZipArchive(_package, ZipArchiveMode.Read, leaveOpen: true);
        return LoadPart(zip, partName);
    }

    /// <summary>A picture part's bytes, read on demand when the host asks for that picture (never kept); null when absent.</summary>
    public byte[]? ReadPartBytes(string partName)
    {
        using var zip = new ZipArchive(_package, ZipArchiveMode.Read, leaveOpen: true);
        ZipArchiveEntry? entry = zip.GetEntry(partName);
        if (entry is null) return null;
        using Stream data = entry.Open();
        var bytes = new MemoryStream(entry.Length is > 0 and < int.MaxValue ? (int)entry.Length : 0);
        data.CopyTo(bytes);
        return bytes.ToArray();
    }

    /// <summary>settings.xml w:compat options the renderer follows.</summary>
    public DocxCompatibilitySettings Compatibility { get; }
    /// <summary>theme1.xml fonts and colors with settings.xml's clrSchemeMapping / themeFontLang; always present (empty without a theme part).</summary>
    public DocxTheme Theme { get; }
    /// <summary>Where the main document, notes, styles… live — from the package relationships.</summary>
    public DocxPackagePartNames PartNames { get; }

    /// <summary>
    /// numbering.xml is read on first use (the first list paragraph counted), and then only the lists asked for are
    /// parsed (<see cref="DocxNumberingDefinitions"/>'s lazy source) — a big document's part is a megabyte for hundreds
    /// of lists, of which a window uses a few; parsing it whole cost the first window ~75 ms.
    /// </summary>
    public DocxNumberingDefinitions Numbering => _numbering ??= LoadNumbering();

    private readonly Stream _package;
    private DocxNumberingDefinitions? _numbering;

    public DocxDocumentParts(Stream package)
    {
        _package = package;
        PartNames = DocxPackagePartNames.Resolve(package);
        using var zip = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        XDocument? settings = LoadPart(zip, PartNames.Settings);          // parsed once: theme mapping, compatibility (note numbering is per section: Word ignores settings.xml footnotePr / endnotePr)
        Theme = new DocxTheme(LoadPart(zip, PartNames.Theme), settings);
        Compatibility = new DocxCompatibilitySettings(settings);
        Styles = new DocxStyleCascade(LoadStyles(zip, PartNames.Styles), Theme) { Compatibility = Compatibility };
        HyperlinkTargets = LoadHyperlinkTargets(zip, DocxPackagePartNames.RelationshipsPartName(PartNames.Document));
        FootnoteHyperlinkTargets = PartNames.Footnotes is { } footnotes ? LoadHyperlinkTargets(zip, DocxPackagePartNames.RelationshipsPartName(footnotes)) : new Dictionary<string, string>();
        EndnoteHyperlinkTargets = PartNames.Endnotes is { } endnotes ? LoadHyperlinkTargets(zip, DocxPackagePartNames.RelationshipsPartName(endnotes)) : new Dictionary<string, string>();
        ImageParts = LoadImageParts(zip, PartNames.Document);
        FootnoteImageParts = LoadImageParts(zip, PartNames.Footnotes);
        EndnoteImageParts = LoadImageParts(zip, PartNames.Endnotes);
        RelatedParts = LoadRelatedParts(zip, PartNames.Document);
        FootnoteRelatedParts = LoadRelatedParts(zip, PartNames.Footnotes);
        EndnoteRelatedParts = LoadRelatedParts(zip, PartNames.Endnotes);
    }

    private DocxNumberingDefinitions LoadNumbering()
    {
        using var zip = new ZipArchive(_package, ZipArchiveMode.Read, leaveOpen: true);
        ZipArchiveEntry? entry = PartNames.Numbering is null ? null : zip.GetEntry(PartNames.Numbering);
        byte[]? bytes = null;
        if (entry is not null)
        {
            using Stream stream = entry.Open();
            var buffer = new MemoryStream(entry.Length > 0 && entry.Length < int.MaxValue ? (int)entry.Length : 0);
            stream.CopyTo(buffer);
            bytes = buffer.ToArray();
        }
        return new DocxNumberingDefinitions(bytes, styleId => Styles.NumberingOf(styleId).numId);
    }

    /// <summary>The scanner's hook: paragraph style / direct numPr → the next label of that list.</summary>
    public DocxListLabelResolver CreateListLabelResolver() => CreateListLabeller().Resolve;

    /// <summary>The list numbering of one pass: the resolver plus its counter state (stored per checkpoint by the compact index).</summary>
    public DocxListLabeller CreateListLabeller() => new(Numbering, Styles, PartNames.Numbering, IsShowableNumberingPicture);

    private Dictionary<string, (string Type, string PartName)>? _numberingRelationships;

    /// <summary>A numbering.xml r:id that names an image a browser shows (the same test DocxWindowedDocument.PartPictureId makes).</summary>
    private bool IsShowableNumberingPicture(string relationshipId)
    {
        Dictionary<string, (string Type, string PartName)> relationships = _numberingRelationships ??= ReadRelationshipsOf(PartNames.Numbering);
        return relationships.TryGetValue(relationshipId, out var target) && target.Type.EndsWith("/image", StringComparison.OrdinalIgnoreCase)
               && WordPictureHtml.ContentTypeOf(target.PartName) is not null;
    }

    /// <summary>
    /// Counts list labels along one pass over body-level paragraphs (<see cref="Resolve"/> is the scanner's hook), and
    /// exports / restores the counter state so a later pass can start counting from a checkpoint instead of block 0.
    /// </summary>
    public sealed class DocxListLabeller
    {
        private readonly DocxNumberingDefinitions _numbering;
        private readonly DocxStyleCascade _styles;
        private readonly DocxListCounter? _counter;
        // keyed by the definition object (one per level, never rebuilt): a record's value hash would hash its strings per paragraph
        private readonly Dictionary<DocxNumberingDefinitions.LevelDefinition, string> _levelCss = new(LevelDefinitionReference.Instance);
        private readonly Dictionary<DocxNumberingDefinitions.LevelDefinition, string> _levelParagraphXml = new(LevelDefinitionReference.Instance);   // serialising the level pPr per paragraph cost more than the scan

        private sealed class LevelDefinitionReference : IEqualityComparer<DocxNumberingDefinitions.LevelDefinition>
        {
            public static readonly LevelDefinitionReference Instance = new();
            public bool Equals(DocxNumberingDefinitions.LevelDefinition? x, DocxNumberingDefinitions.LevelDefinition? y) => ReferenceEquals(x, y);
            public int GetHashCode(DocxNumberingDefinitions.LevelDefinition definition) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(definition);
        }

        private readonly string? _numberingPartName;

        private readonly Func<string, bool>? _isShowablePicture;

        internal DocxListLabeller(DocxNumberingDefinitions numbering, DocxStyleCascade styles, string? numberingPartName = null, Func<string, bool>? isShowablePicture = null)
        {
            _numberingPartName = numberingPartName;
            _isShowablePicture = isShowablePicture;
            _numbering = numbering;
            _styles = styles;
            _counter = numbering.IsEmpty ? null : new DocxListCounter(numbering);
        }

        public DocxListLabel? Resolve(string? styleId, int numId, int level)
        {
            if (_counter is null) return null;
            if (numId < 0 && styleId is not null)
                (numId, level) = _styles.NumberingOf(styleId);
            if (numId <= 0 || _counter.NextWithDefinition(numId, level) is not { } next) return null;
            DocxNumberingDefinitions.LevelDefinition? definition = next.definition;
            string label = next.label;
            bool isBulletMapped = false;
            // w:lvlPicBulletId (17.9.10): the picture of a numPicBullet, read through the numbering part's relationships (the "x:"
            // picture id); a missing picture or one a browser cannot show keeps the level's text bullet (resolved below)
            if (definition is { PictureBulletId: { } pictureBulletId } && _numberingPartName is not null
                && _numbering.PictureBulletOf(pictureBulletId) is { } picture && _isShowablePicture?.Invoke(picture.RelationshipId) == true)
                label = WordListLabelHtml.PictureLabel("x:" + _numberingPartName + "#" + picture.RelationshipId,
                                                       picture.WidthPt > 0 ? picture.WidthPt : 1, picture.HeightPt > 0 ? picture.HeightPt : 1);
            else if (definition is { Format: "bullet" })
            {
                (label, bool keepsFont) = WordSymbolFontBullet.Resolve(label, DocxNumberingDefinitions.FontOf(definition.RunProperties));
                isBulletMapped = !keepsFont;
            }
            string css = "";
            if (definition is not null && !_levelCss.TryGetValue(definition, out css!))
                _levelCss[definition] = css = _styles.LevelRunCss(definition.RunProperties, isBulletMapped);
            string paragraphXml = "";
            if (definition is not null && !_levelParagraphXml.TryGetValue(definition, out paragraphXml!))
                _levelParagraphXml[definition] = paragraphXml = definition.ParagraphProperties?.ToString(SaveOptions.DisableFormatting) ?? "";
            return new DocxListLabel(label, DocxListLabel.SuffixOf(definition?.Suffix), css, paragraphXml, DocxListLabel.AlignmentOf(definition?.Justification),
                                     string.Join("\u0001", _counter.ContextOf(numId, level)), definition?.Legacy);
        }

        /// <summary>The counter state now (empty when the document has no numbering).</summary>
        public byte[] ExportState() => _counter?.ExportState() ?? [];

        public void ImportState(byte[] state)
        {
            _counter?.ImportState(state);
        }
    }

    /// <summary>
    /// styles.xml through a forward-only reader keeping only w:docDefaults and the w:style elements:
    /// w:latentStyles (often most of the part — hundreds of built-in style switches) is skipped
    /// without building nodes. Same XDocument shape for DocxStyleCascade, a fraction of the cost.
    /// </summary>
    private static XDocument? LoadStyles(ZipArchive zip, string? name)
    {
        try { return ReadStyles(zip, name); }
        catch (XmlException)
        {
            // malformed: what precedes the first error, the same elements the streaming path keeps
            XDocument? whole = WordDocumentXml.LoadPartKeepingContentBeforeFirstError(() => zip.GetEntry(name!)!.Open());
            if (whole?.Root is not { } root || root.Name != DocxStyleCascade.W + "styles") return null;
            root.DescendantNodes().OfType<XComment>().Remove();   // the streaming path ignores comments
            return new XDocument(new XElement(root.Name, root.Elements().Where(element => element.Name == DocxStyleCascade.W + "docDefaults" || element.Name == DocxStyleCascade.W + "style")));
        }
    }

    private static XDocument? ReadStyles(ZipArchive zip, string? name)
    {
        ZipArchiveEntry? entry = name is null ? null : zip.GetEntry(name);
        if (entry is null) return null;
        using Stream stream = entry.Open();
        using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = true, IgnoreComments = true, DtdProcessing = DtdProcessing.Ignore });
        if (!reader.ReadToFollowing("styles", DocxStyleCascade.W.NamespaceName)) return null;
        var root = new XElement(DocxStyleCascade.W + "styles");
        if (reader.IsEmptyElement) return new XDocument(root);
        reader.Read();
        while (!reader.EOF && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == 0))
        {
            if (reader.NodeType == XmlNodeType.Element && reader.Depth == 1
                && reader.NamespaceURI == DocxStyleCascade.W.NamespaceName && reader.LocalName is "docDefaults" or "style")
                root.Add(XNode.ReadFrom(reader));                 // advances past the element
            else if (reader.NodeType == XmlNodeType.Element && reader.Depth == 1) reader.Skip();   // latentStyles and anything else
            else reader.Read();
        }
        return new XDocument(root);
    }

    private static XDocument? LoadPart(ZipArchive zip, string? name)
    {
        ZipArchiveEntry? entry = name is null ? null : zip.GetEntry(name);
        if (entry is null) return null;
        return WordDocumentXml.LoadPartKeepingContentBeforeFirstError(entry.Open);
    }
}

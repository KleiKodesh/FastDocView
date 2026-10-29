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
    public DocxNoteNumberFormats NoteNumberFormats { get; }
    /// <summary>The main part's hyperlink relationships: r:id → target (external URLs as written). Small; read at open.</summary>
    public IReadOnlyDictionary<string, string> HyperlinkTargets { get; }

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

    /// <summary>settings.xml w:compat options the renderer follows.</summary>
    public DocxCompatibilitySettings Compatibility { get; }
    /// <summary>theme1.xml fonts and colors with settings.xml's clrSchemeMapping / themeFontLang; always present (empty without a theme part).</summary>
    public DocxTheme Theme { get; }
    /// <summary>Where the main document, notes, styles… live — from the package relationships.</summary>
    public DocxPackagePartNames PartNames { get; }

    /// <summary>
    /// numbering.xml is only needed by the index pass (list labels are stored in the index), so it
    /// is parsed on first use — a reopen from a cached index never pays for it (~75 ms on big files).
    /// </summary>
    public DocxNumberingDefinitions Numbering => _numbering ??= LoadNumbering();

    private readonly Stream _package;
    private DocxNumberingDefinitions? _numbering;

    public DocxDocumentParts(Stream package)
    {
        _package = package;
        PartNames = DocxPackagePartNames.Resolve(package);
        using var zip = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        XDocument? settings = LoadPart(zip, PartNames.Settings);          // parsed once: theme mapping, note formats, compatibility
        Theme = new DocxTheme(LoadPart(zip, PartNames.Theme), settings);
        Compatibility = new DocxCompatibilitySettings(settings);
        Styles = new DocxStyleCascade(LoadStyles(zip, PartNames.Styles), Theme) { Compatibility = Compatibility };
        NoteNumberFormats = new DocxNoteNumberFormats(settings);
        HyperlinkTargets = LoadHyperlinkTargets(zip, DocxPackagePartNames.RelationshipsPartName(PartNames.Document));
    }

    private DocxNumberingDefinitions LoadNumbering()
    {
        using var zip = new ZipArchive(_package, ZipArchiveMode.Read, leaveOpen: true);
        return new DocxNumberingDefinitions(LoadPart(zip, PartNames.Numbering), styleId => Styles.NumberingOf(styleId).numId);
    }

    /// <summary>The scanner's hook: paragraph style / direct numPr → the next label of that list.</summary>
    public DocxListLabelResolver CreateListLabelResolver()
    {
        if (Numbering.IsEmpty) return (_, _, _) => null;
        var counter = new DocxListCounter(Numbering);
        var levelCss = new Dictionary<DocxNumberingDefinitions.LevelDefinition, string>();
        return (styleId, numId, level) =>
        {
            if (numId < 0 && styleId is not null)
                (numId, level) = Styles.NumberingOf(styleId);
            if (numId <= 0 || counter.NextWithDefinition(numId, level) is not { } next) return null;
            DocxNumberingDefinitions.LevelDefinition? definition = next.definition;
            string css = "";
            if (definition is not null && !levelCss.TryGetValue(definition, out css!))
                levelCss[definition] = css = Styles.LevelRunCss(definition.RunProperties, isBulletMapped: definition.Format == "bullet" && next.label != definition.Text);
            return new DocxListLabel(next.label, DocxListLabel.SuffixOf(definition?.Suffix), css,
                                     definition?.ParagraphProperties?.ToString(SaveOptions.DisableFormatting) ?? "", DocxListLabel.AlignmentOf(definition?.Justification));
        };
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

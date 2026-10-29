using System.Xml;
using SeekableDeflate;

namespace FastDocView.Docx;

/// <summary>
/// Where a .docx keeps its parts. The main document is the target of the package's officeDocument
/// relationship (_rels/.rels), and styles / numbering / settings / theme / footnotes / endnotes are
/// targets of the main part's own relationships — Word names them word/document.xml, word/styles.xml…
/// but other writers do not (LibreOffice's own test file tdf104713_undefinedStyles.docx keeps its
/// body in word/trial.xml). Resolved the way LibreOffice does (oox/source/core/relationshandler.cxx
/// lclGetRelationsPath: "dir/_rels/file.xml.rels", the package root's "_rels/.rels";
/// oox/source/core/relations.cxx getFragmentPathFromRelation: a leading '/' is absolute, '..'
/// drops a folder, everything else is relative to the source part's folder; types match the
/// transitional or the strict namespace, ignoring ASCII case). A package without relationship
/// parts falls back to Word's conventional names.
/// </summary>
public sealed record DocxPackagePartNames(string Document, string? Footnotes, string? Endnotes, string? Styles,
                                          string? Numbering, string? Settings, string? Theme)
{
    private const string TransitionalRelationshipPrefix = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    private const string StrictRelationshipPrefix = "http://purl.oclc.org/ooxml/officeDocument/relationships/";
    private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";

    public static readonly DocxPackagePartNames WordConventional = new("word/document.xml", "word/footnotes.xml", "word/endnotes.xml",
        "word/styles.xml", "word/numbering.xml", "word/settings.xml", "word/theme/theme1.xml");

    public static DocxPackagePartNames Resolve(Stream package)
    {
        List<(string type, string target)>? packageRelationships = ReadRelationships(package, "_rels/.rels", sourcePartName: "");
        string? document = packageRelationships is null ? null : TargetOf(packageRelationships, "officeDocument");
        if (document is null) return WordConventional;
        List<(string type, string target)> partRelationships = ReadRelationships(package, RelationshipsPartName(document), document) ?? [];
        return new DocxPackagePartNames(document,
            TargetOf(partRelationships, "footnotes"), TargetOf(partRelationships, "endnotes"), TargetOf(partRelationships, "styles"),
            TargetOf(partRelationships, "numbering"), TargetOf(partRelationships, "settings"), TargetOf(partRelationships, "theme"));
    }

    private static string? TargetOf(List<(string type, string target)> relationships, string typeName)
    {
        foreach ((string type, string target) in relationships)
            if (type.Equals(TransitionalRelationshipPrefix + typeName, StringComparison.OrdinalIgnoreCase)
                || type.Equals(StrictRelationshipPrefix + typeName, StringComparison.OrdinalIgnoreCase))
                return target;
        return null;
    }

    /// <summary>"word/document.xml" → "word/_rels/document.xml.rels".</summary>
    internal static string RelationshipsPartName(string partName)
    {
        int slash = partName.LastIndexOf('/');
        return partName.Substring(0, slash + 1) + "_rels/" + partName.Substring(slash + 1) + ".rels";
    }

    /// <summary>Internal relationships of a part with their targets resolved to zip entry names, or null when there is no such relationships part.</summary>
    private static List<(string type, string target)>? ReadRelationships(Stream package, string relationshipsPartName, string sourcePartName)
    {
        ZipEntryLocation? entry = ZipEntryLocator.Locate(package, relationshipsPartName);
        if (entry is null) return null;
        if (!entry.IsStored && !entry.IsDeflate) throw new InvalidDataException($"Unsupported zip compression method {entry.CompressionMethod} for {relationshipsPartName}.");
        var relationships = new List<(string, string)>();
        using Stream view = entry.IsDeflate
            ? new System.IO.Compression.DeflateStream(new ZipStoredEntryStream(package, entry.DataOffset, entry.CompressedSize), System.IO.Compression.CompressionMode.Decompress)
            : new ZipStoredEntryStream(package, entry.DataOffset, entry.CompressedSize);
        try
        {
            using var reader = XmlReader.Create(view, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore, CloseInput = false });
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "Relationship" || reader.NamespaceURI != RelationshipsNamespace) continue;
                string? type = reader.GetAttribute("Type"), target = reader.GetAttribute("Target");
                if (type is null || string.IsNullOrEmpty(target) || reader.GetAttribute("TargetMode") == "External") continue;
                relationships.Add((type, ResolveTarget(sourcePartName, target!)));
            }
        }
        catch (XmlException exception) { throw new InvalidDataException($"{relationshipsPartName} is not well-formed XML: {exception.Message}", exception); }
        return relationships;
    }

    private static string ResolveTarget(string sourcePartName, string target)
    {
        target = Uri.UnescapeDataString(target);
        if (target[0] == '/') return target.Substring(1);
        var segments = new List<string>(sourcePartName.Split('/'));
        segments.RemoveAt(segments.Count - 1);                 // the source part's own file name
        foreach (string segment in target.Split('/'))
        {
            if (segment is "" or ".") continue;                 // "./styles.xml", "a//b": same folder (URI dot-segment removal, RFC 3986 §5.2.4)
            if (segment == "..") { if (segments.Count > 0) segments.RemoveAt(segments.Count - 1); }
            else segments.Add(segment);
        }
        return string.Join("/", segments);
    }
}

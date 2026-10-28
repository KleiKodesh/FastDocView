using System.IO.Compression;
using System.Xml.Linq;

namespace WordDocumentRandomAccess.DocxHtml;

/// <summary>
/// footnotes.xml / endnotes.xml bodies by id, and the number format of each kind from
/// settings.xml (w:footnotePr / w:endnotePr; Word's defaults are decimal and lowerRoman).
/// The note parts are small next to the body (hundreds of KB), so they are loaded whole; the
/// separator notes Word keeps there (w:type="separator" / "continuationSeparator") are skipped.
/// </summary>
public sealed class DocxNoteBodies
{
    private static readonly XNamespace W = DocxStyleCascade.W;

    private readonly Dictionary<int, XElement> _footnotes;
    private readonly Dictionary<int, XElement> _endnotes;

    public string FootnoteNumberFormat { get; } = "decimal";
    public string EndnoteNumberFormat { get; } = "lowerRoman";

    public DocxNoteBodies(ZipArchive package)
    {
        _footnotes = LoadNotes(package, "word/footnotes.xml", "footnote");
        _endnotes = LoadNotes(package, "word/endnotes.xml", "endnote");
        XElement? settings = LoadPart(package, "word/settings.xml")?.Root;
        FootnoteNumberFormat = (string?)settings?.Element(W + "footnotePr")?.Element(W + "numFmt")?.Attribute(W + "val") ?? FootnoteNumberFormat;
        EndnoteNumberFormat = (string?)settings?.Element(W + "endnotePr")?.Element(W + "numFmt")?.Attribute(W + "val") ?? EndnoteNumberFormat;
    }

    public XElement? Find(WordDocumentNoteKind kind, int noteId)
        => (kind == WordDocumentNoteKind.Footnote ? _footnotes : _endnotes).TryGetValue(noteId, out XElement? note) ? note : null;

    public string FormatNumber(WordDocumentNoteKind kind, int number)
        => DocxNumberingDefinitions.FormatNumber(number, kind == WordDocumentNoteKind.Footnote ? FootnoteNumberFormat : EndnoteNumberFormat);

    private static Dictionary<int, XElement> LoadNotes(ZipArchive package, string partName, string elementName)
    {
        var notes = new Dictionary<int, XElement>();
        XElement? root = LoadPart(package, partName)?.Root;
        if (root is null) return notes;
        foreach (XElement note in root.Elements(W + elementName))
        {
            string? type = (string?)note.Attribute(W + "type");
            if (type is "separator" or "continuationSeparator" or "continuationNotice") continue;
            if (int.TryParse((string?)note.Attribute(W + "id"), out int id)) notes[id] = note;
        }
        return notes;
    }

    private static XDocument? LoadPart(ZipArchive package, string name)
    {
        ZipArchiveEntry? entry = package.GetEntry(name);
        if (entry is null) return null;
        using Stream stream = entry.Open();
        return XDocument.Load(stream);
    }
}

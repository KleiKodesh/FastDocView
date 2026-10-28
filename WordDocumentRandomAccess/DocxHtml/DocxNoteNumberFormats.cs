using System.IO.Compression;
using System.Xml.Linq;

namespace WordDocumentRandomAccess.DocxHtml;

/// <summary>
/// How footnotes and endnotes are numbered: settings.xml w:footnotePr / w:endnotePr numFmt
/// (Word's defaults: decimal and lowerRoman). The note bodies themselves are not here — they are
/// read one at a time through <see cref="Docx.DocxNoteXmlReader"/>.
/// </summary>
public sealed class DocxNoteNumberFormats
{
    private static readonly XNamespace W = DocxStyleCascade.W;

    public string FootnoteNumberFormat { get; } = "decimal";
    public string EndnoteNumberFormat { get; } = "lowerRoman";

    public DocxNoteNumberFormats(ZipArchive package)
    {
        ZipArchiveEntry? entry = package.GetEntry("word/settings.xml");
        if (entry is null) return;
        XElement? settings;
        using (Stream stream = entry.Open()) settings = XDocument.Load(stream).Root;
        FootnoteNumberFormat = (string?)settings?.Element(W + "footnotePr")?.Element(W + "numFmt")?.Attribute(W + "val") ?? FootnoteNumberFormat;
        EndnoteNumberFormat = (string?)settings?.Element(W + "endnotePr")?.Element(W + "numFmt")?.Attribute(W + "val") ?? EndnoteNumberFormat;
    }

    public string Format(WordDocumentNoteKind kind, int number)
        => DocxNumberingDefinitions.FormatNumber(number, kind == WordDocumentNoteKind.Footnote ? FootnoteNumberFormat : EndnoteNumberFormat);
}

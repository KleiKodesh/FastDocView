using System.IO.Compression;
using System.Xml.Linq;
using WordDocumentRandomAccess.DocxHtml;

namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// The small parts of a .docx that are loaded whole (styles, numbering) — a few hundred KB at
/// most, unlike the body. Read through ZipArchive over any seekable stream, so it works the same
/// on a file-system host and in the browser.
/// </summary>
public sealed class DocxDocumentParts
{
    public DocxStyleCascade Styles { get; }
    public DocxNumberingDefinitions Numbering { get; }

    public DocxDocumentParts(Stream package)
    {
        using var zip = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        Styles = new DocxStyleCascade(LoadPart(zip, "word/styles.xml"));
        Numbering = new DocxNumberingDefinitions(LoadPart(zip, "word/numbering.xml"));
    }

    /// <summary>The scanner's hook: paragraph style / direct numPr → the next label of that list.</summary>
    public DocxListLabelResolver CreateListLabelResolver()
    {
        if (Numbering.IsEmpty) return (_, _, _) => null;
        var counter = new DocxListCounter(Numbering);
        return (styleId, numId, level) =>
        {
            if (numId < 0 && styleId is not null)
                (numId, level) = Styles.NumberingOf(styleId);
            return numId > 0 ? counter.Next(numId, level) : null;
        };
    }

    private static XDocument? LoadPart(ZipArchive zip, string name)
    {
        ZipArchiveEntry? entry = zip.GetEntry(name);
        if (entry is null) return null;
        using Stream stream = entry.Open();
        return XDocument.Load(stream);
    }
}

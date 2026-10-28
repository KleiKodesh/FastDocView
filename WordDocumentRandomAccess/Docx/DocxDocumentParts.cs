using System.IO.Compression;
using System.Xml.Linq;
using WordDocumentRandomAccess.DocxHtml;

namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// The small parts of a .docx that are loaded whole (styles, theme, settings, numbering) — a few hundred KB at
/// most, unlike the body. Read through ZipArchive over any seekable stream, so it works the same
/// on a file-system host and in the browser.
/// </summary>
public sealed class DocxDocumentParts
{
    public DocxStyleCascade Styles { get; }
    public DocxNoteNumberFormats NoteNumberFormats { get; }
    /// <summary>theme1.xml fonts and colors with settings.xml's clrSchemeMapping / themeFontLang; always present (empty without a theme part).</summary>
    public DocxTheme Theme { get; }

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
        using var zip = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        Theme = new DocxTheme(LoadPart(zip, ThemePartName(zip)), LoadPart(zip, "word/settings.xml"));
        Styles = new DocxStyleCascade(LoadPart(zip, "word/styles.xml"), Theme);
        NoteNumberFormats = new DocxNoteNumberFormats(zip);
    }

    private DocxNumberingDefinitions LoadNumbering()
    {
        using var zip = new ZipArchive(_package, ZipArchiveMode.Read, leaveOpen: true);
        return new DocxNumberingDefinitions(LoadPart(zip, "word/numbering.xml"));
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

    /// <summary>Word writes word/theme/theme1.xml; any other single theme part under word/theme/ is accepted too.</summary>
    private static string ThemePartName(ZipArchive zip)
    {
        if (zip.GetEntry("word/theme/theme1.xml") is not null) return "word/theme/theme1.xml";
        foreach (ZipArchiveEntry entry in zip.Entries)
            if (entry.FullName.StartsWith("word/theme/", StringComparison.OrdinalIgnoreCase) && entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                return entry.FullName;
        return "word/theme/theme1.xml";
    }

    private static XDocument? LoadPart(ZipArchive zip, string name)
    {
        ZipArchiveEntry? entry = zip.GetEntry(name);
        if (entry is null) return null;
        using Stream stream = entry.Open();
        return XDocument.Load(stream);
    }
}

using System.Globalization;
using System.Xml.Linq;

namespace FastDocView.Docx;

/// <summary>
/// <see cref="WordDocumentProperties"/> of a .docx (and a Flat OPC file, whose virtual package carries the same parts): the
/// core properties part docProps/core.xml (ISO/IEC 29500-2 11: dc:title, dc:subject, dc:creator, cp:keywords,
/// dc:description, cp:lastModifiedBy, cp:revision, dcterms:created, dcterms:modified, cp:lastPrinted, cp:category,
/// cp:contentStatus) and the extended properties part docProps/app.xml (ISO/IEC 29500-1 22.2: Pages, Words, Characters,
/// CharactersWithSpaces, Lines, Paragraphs, TotalTime in minutes, Application, Company, Manager, Template). Element to
/// property as LibreOffice imports them (oox/source/docprop/docprophandler.cxx: Characters = the non-whitespace count,
/// CharactersWithSpaces = the character count; TotalTime minutes). Two small parts, read only when asked.
/// </summary>
internal static class DocxDocumentProperties
{
    private static readonly XNamespace DublinCore = "http://purl.org/dc/elements/1.1/";
    private static readonly XNamespace DublinCoreTerms = "http://purl.org/dc/terms/";
    private static readonly XNamespace CoreProperties = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private static readonly XNamespace ExtendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

    public static WordDocumentProperties Read(Func<string, XDocument?> readPart)
    {
        XElement? core = readPart("docProps/core.xml")?.Root, app = readPart("docProps/app.xml")?.Root;
        if (core is null && app is null) return WordDocumentProperties.None;
        string? Core(XNamespace ns, string name) => WordDocumentProperties.Text(core?.Element(ns + name)?.Value);
        string? App(string name) => WordDocumentProperties.Text(app?.Element(ExtendedProperties + name)?.Value);
        return new WordDocumentProperties
        {
            Title = Core(DublinCore, "title"),
            Subject = Core(DublinCore, "subject"),
            Author = Core(DublinCore, "creator"),
            Keywords = Core(CoreProperties, "keywords"),
            Comments = Core(DublinCore, "description"),
            LastModifiedBy = Core(CoreProperties, "lastModifiedBy"),
            Revision = Core(CoreProperties, "revision"),
            Category = Core(CoreProperties, "category"),
            ContentStatus = Core(CoreProperties, "contentStatus"),
            Created = Date(Core(DublinCoreTerms, "created")),
            Modified = Date(Core(DublinCoreTerms, "modified")),
            LastPrinted = Date(Core(CoreProperties, "lastPrinted")),
            Pages = WordDocumentProperties.Count(App("Pages")),
            Words = WordDocumentProperties.Count(App("Words")),
            Characters = WordDocumentProperties.Count(App("Characters")),
            CharactersWithSpaces = WordDocumentProperties.Count(App("CharactersWithSpaces")),
            Lines = WordDocumentProperties.Count(App("Lines")),
            Paragraphs = WordDocumentProperties.Count(App("Paragraphs")),
            TotalEditingMinutes = WordDocumentProperties.Count(App("TotalTime")),
            Application = App("Application"),
            Company = App("Company"),
            Manager = App("Manager"),
            Template = App("Template"),
        };
    }

    /// <summary>A W3CDTF date (dcterms:W3CDTF, "2026-11-26T10:00:00Z"): UTC when it names a zone, as written otherwise.</summary>
    internal static DateTime? Date(string? value)
    {
        if (value is null) return null;
        bool hasZone = value.EndsWith("Z", StringComparison.OrdinalIgnoreCase) || (value.Length > 19 && value.LastIndexOfAny(['+', '-']) > 18);
        if (hasZone && DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out DateTimeOffset moment))
            return moment.UtcDateTime;
        return DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime wallTime)
            ? DateTime.SpecifyKind(wallTime, DateTimeKind.Unspecified) : null;
    }
}

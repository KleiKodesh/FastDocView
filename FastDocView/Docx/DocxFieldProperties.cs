using System.Globalization;
using System.Xml.Linq;
using FastDocView.Fields;

namespace FastDocView.Docx;

/// <summary>
/// The document properties .docx fields show (AUTHOR, TITLE, CREATEDATE, NUMWORDS, DOCPROPERTY, DOCVARIABLE ...), read once
/// on first use from the small property parts: docProps/core.xml (ECMA-376 Part 2 11, Dublin Core + cp:), docProps/app.xml
/// (Part 1 22.2 extended properties), docProps/custom.xml (22.3 custom properties) and settings.xml w:docVars (Part 1
/// 17.15.1.31). Field name → element as LibreOffice maps them (sw/source/writerfilter/dmapper/DomainMapper_Impl.cxx
/// FIELD_AUTHOR / TITLE / SUBJECT / KEYWORDS / COMMENTS / LASTSAVEDBY → the matching document property).
/// </summary>
internal sealed class DocxFieldProperties
{
    private static readonly XNamespace DublinCore = "http://purl.org/dc/elements/1.1/";
    private static readonly XNamespace DublinCoreTerms = "http://purl.org/dc/terms/";
    private static readonly XNamespace CoreProperties = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private static readonly XNamespace ExtendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
    private static readonly XNamespace CustomProperties = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private readonly Dictionary<string, string> _properties = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, DateTime> _dates = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _variables = new(StringComparer.Ordinal);

    public DocxFieldProperties(Func<string, XDocument?> readPart, string? settingsPartName)
    {
        if (readPart("docProps/core.xml")?.Root is { } core)
        {
            Add("TITLE", core.Element(DublinCore + "title"));
            Add("SUBJECT", core.Element(DublinCore + "subject"));
            Add("AUTHOR", core.Element(DublinCore + "creator"));
            Add("KEYWORDS", core.Element(CoreProperties + "keywords"));
            Add("COMMENTS", core.Element(DublinCore + "description"));
            Add("LASTSAVEDBY", core.Element(CoreProperties + "lastModifiedBy"));
            Add("REVNUM", core.Element(CoreProperties + "revision"));
            Add("CATEGORY", core.Element(CoreProperties + "category"));
            AddDate("CREATEDATE", core.Element(DublinCoreTerms + "created"));
            AddDate("SAVEDATE", core.Element(DublinCoreTerms + "modified"));
            AddDate("PRINTDATE", core.Element(CoreProperties + "lastPrinted"));
        }
        if (readPart("docProps/app.xml")?.Root is { } app)
        {
            Add("TEMPLATE", app.Element(ExtendedProperties + "Template"));
            Add("NUMWORDS", app.Element(ExtendedProperties + "Words"));
            Add("NUMCHARS", app.Element(ExtendedProperties + "Characters"));
            Add("EDITTIME", app.Element(ExtendedProperties + "TotalTime"));
            Add("COMPANY", app.Element(ExtendedProperties + "Company"));
            Add("MANAGER", app.Element(ExtendedProperties + "Manager"));
            Add("CHARACTERSWITHSPACES", app.Element(ExtendedProperties + "CharactersWithSpaces"));
            Add("LINES", app.Element(ExtendedProperties + "Lines"));
            Add("PARAGRAPHS", app.Element(ExtendedProperties + "Paragraphs"));
            Add("HYPERLINKBASE", app.Element(ExtendedProperties + "HyperlinkBase"));
        }
        if (readPart("docProps/custom.xml")?.Root is { } custom)
            foreach (XElement property in custom.Elements(CustomProperties + "property"))
                if ((string?)property.Attribute("name") is { Length: > 0 } name && property.Elements().FirstOrDefault() is { } value)
                    _properties[name] = value.Value;
        if (settingsPartName is not null && readPart(settingsPartName)?.Root?.Element(W + "docVars") is { } variables)
            foreach (XElement variable in variables.Elements(W + "docVar"))
                if ((string?)variable.Attribute(W + "name") is { Length: > 0 } name) _variables[name] = (string?)variable.Attribute(W + "val") ?? "";
    }

    private void Add(string name, XElement? element)
    {
        if (element is not null) _properties[name] = element.Value;
    }

    private void AddDate(string name, XElement? element)
    {
        if (element is not null && DateTime.TryParse(element.Value, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out DateTime date))
            _dates[name] = date.ToLocalTime();
    }

    public string? Property(string name) => _properties.TryGetValue(name, out string? value) ? value : null;
    public DateTime? Date(string name) => _dates.TryGetValue(name, out DateTime value) ? value : null;
    public string? Variable(string name) => _variables.TryGetValue(name, out string? value) ? value : null;

    /// <summary>A field context over these properties, read on first use.</summary>
    public static WordFieldContext Context(Func<DocxFieldProperties> properties, string filePath, Func<string, string?>? bookmarkText)
    {
        DocxFieldProperties? loaded = null;
        DocxFieldProperties Loaded() => loaded ??= properties();
        long? size = null;
        try { size = File.Exists(filePath) ? new FileInfo(filePath).Length : null; } catch (IOException) { } catch (UnauthorizedAccessException) { }
        return new WordFieldContext
        {
            DocumentProperty = name => Loaded().Property(name),
            DocumentDate = name => Loaded().Date(name),
            DocumentVariable = name => Loaded().Variable(name),
            FileName = Path.GetFileName(filePath),
            FilePath = filePath,
            FileSize = size,
            BookmarkText = bookmarkText,
        };
    }
}

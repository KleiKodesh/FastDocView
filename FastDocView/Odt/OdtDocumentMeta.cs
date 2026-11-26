using System.Globalization;
using System.Xml;

namespace FastDocView.Odt;

/// <summary>
/// The document information an OpenDocument text's fields show (ODF 1.3 part 3 3.3 office:meta: dc:title, dc:subject,
/// dc:description, meta:keyword, meta:initial-creator, meta:creation-date, dc:date, meta:user-defined) and meta:generator.
/// Read from meta.xml (package) or the office:meta before a .fodt's office:body - a small part, never the body.
/// </summary>
internal sealed class OdtDocumentMeta
{
    public static readonly OdtDocumentMeta Empty = new();
    private const string MetaNamespace = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
    private const string DublinCoreNamespace = "http://purl.org/dc/elements/1.1/";

    public string? Generator { get; private set; }
    public string? Title { get; private set; }
    public string? Subject { get; private set; }
    public string? Description { get; private set; }
    public List<string> Keywords { get; } = new();
    public string? InitialCreator { get; private set; }
    public DateTime? CreationDate { get; private set; }
    public DateTime? ModificationDate { get; private set; }
    public Dictionary<string, string> UserDefined { get; } = new(StringComparer.Ordinal);
    /// <summary>meta:document-statistic (ODF 1.3 part 3 4.3.2.18) by attribute local name (word-count, paragraph-count...).</summary>
    public Dictionary<string, string> Statistics { get; } = new(StringComparer.Ordinal);
    // the properties dialog's values (ODF 1.3 part 3: dc:creator 4.3.2.7 the last author, meta:print-date 4.3.2.11,
    // meta:template 4.3.2.12 xlink:title / href, meta:editing-cycles 4.3.2.16, meta:editing-duration 4.3.2.17;
    // the date texts as written, so a zone is kept)
    private string? _creator, _creationDateText, _modificationDateText, _printDateText, _editingCycles, _editingDuration, _template;

    /// <summary>
    /// <see cref="WordDocumentProperties"/> from office:meta, as LibreOffice imports it (sfx2/source/doc/SfxDocumentMetaData.cxx
    /// s_stdMeta: meta:initial-creator = author, dc:creator = modified by, meta:editing-cycles, meta:editing-duration,
    /// meta:template; the statistics by s_stdStatAttrs: meta:page-count, meta:word-count, meta:paragraph-count,
    /// meta:character-count = with spaces, meta:non-whitespace-character-count = without - the pairing
    /// oox/source/docprop/docprophandler.cxx uses for .docx Characters / CharactersWithSpaces; meta:row-count = "the number of
    /// lines" by ODF 1.3 part 3 19.334). ODF has no category, company, manager or status; meta:keyword elements are joined with ", " (LibreOffice's keyword list).
    /// </summary>
    public WordDocumentProperties Document => this == Empty ? WordDocumentProperties.None : new WordDocumentProperties
    {
        Title = WordDocumentProperties.Text(Title),
        Subject = WordDocumentProperties.Text(Subject),
        Comments = WordDocumentProperties.Text(Description),
        Keywords = WordDocumentProperties.Text(string.Join(", ", Keywords.Select(keyword => keyword.Trim()).Where(keyword => keyword.Length > 0))),
        Author = WordDocumentProperties.Text(InitialCreator),
        LastModifiedBy = WordDocumentProperties.Text(_creator),
        Application = WordDocumentProperties.Text(Generator),
        Template = WordDocumentProperties.Text(_template),
        Revision = WordDocumentProperties.Count(_editingCycles)?.ToString(CultureInfo.InvariantCulture),
        TotalEditingMinutes = EditingMinutes(_editingDuration),
        Created = Docx.DocxDocumentProperties.Date(WordDocumentProperties.Text(_creationDateText)),
        Modified = Docx.DocxDocumentProperties.Date(WordDocumentProperties.Text(_modificationDateText)),
        LastPrinted = Docx.DocxDocumentProperties.Date(WordDocumentProperties.Text(_printDateText)),
        Pages = Statistic("page-count"),
        Words = Statistic("word-count"),
        Paragraphs = Statistic("paragraph-count"),
        CharactersWithSpaces = Statistic("character-count"),
        Characters = Statistic("non-whitespace-character-count"),
        Lines = Statistic("row-count"),
    };

    private int? Statistic(string name) => Statistics.TryGetValue(name, out string? value) ? WordDocumentProperties.Count(value) : null;

    /// <summary>An xsd:duration (PT1H2M3S, P1DT2H) in whole minutes.</summary>
    private static int? EditingMinutes(string? value)
    {
        if (WordDocumentProperties.Text(value) is not { } text) return null;
        try
        {
            TimeSpan span = XmlConvert.ToTimeSpan(text);
            return span >= TimeSpan.Zero && span.TotalMinutes <= int.MaxValue ? (int)span.TotalMinutes : null;
        }
        catch (Exception exception) when (exception is FormatException or OverflowException) { return null; }
    }

    /// <summary>Reads office:meta's children from <paramref name="reader"/> until office:body / the styles begin (null when it never reaches meta).</summary>
    public static OdtDocumentMeta Read(XmlReader reader, string officeNamespace)
    {
        var meta = new OdtDocumentMeta();
        reader.Read();
        while (!reader.EOF)
        {
            if (reader.NodeType != XmlNodeType.Element) { reader.Read(); continue; }
            if (reader.NamespaceURI == officeNamespace && reader.LocalName is "body" or "styles" or "automatic-styles" or "master-styles" or "font-face-decls" or "settings" or "scripts") break;
            bool isMeta = reader.NamespaceURI == MetaNamespace, isDublinCore = reader.NamespaceURI == DublinCoreNamespace;
            if (isMeta && reader.LocalName == "document-statistic")
            {
                if (reader.MoveToFirstAttribute())
                {
                    do if (reader.NamespaceURI == MetaNamespace) meta.Statistics[reader.LocalName] = reader.Value;
                    while (reader.MoveToNextAttribute());
                    reader.MoveToElement();
                }
                reader.Read();
                continue;
            }
            if (isMeta && reader.LocalName == "template")
            {
                const string XLinkNamespace = "http://www.w3.org/1999/xlink";
                meta._template = reader.GetAttribute("title", XLinkNamespace) is { Length: > 0 } title ? title : reader.GetAttribute("href", XLinkNamespace);
                reader.Skip();
                continue;
            }
            if ((!isMeta && !isDublinCore) || reader.IsEmptyElement && !(isMeta && reader.LocalName == "user-defined")) { reader.Read(); continue; }
            string local = reader.LocalName;
            string? name = isMeta && local == "user-defined" ? reader.GetAttribute("name", MetaNamespace) : null;
            string value = reader.ReadElementContentAsString();
            switch (isMeta ? "meta:" + local : "dc:" + local)
            {
                case "meta:generator": meta.Generator = value; break;
                case "dc:title": meta.Title = value; break;
                case "dc:subject": meta.Subject = value; break;
                case "dc:description": meta.Description = value; break;
                case "meta:keyword": meta.Keywords.Add(value); break;
                case "meta:initial-creator": meta.InitialCreator = value; break;
                case "meta:creation-date": meta.CreationDate = ParseDate(value); meta._creationDateText = value; break;
                case "dc:date": meta.ModificationDate = ParseDate(value); meta._modificationDateText = value; break;
                case "dc:creator": meta._creator = value; break;
                case "meta:print-date": meta._printDateText = value; break;
                case "meta:editing-cycles": meta._editingCycles = value; break;
                case "meta:editing-duration": meta._editingDuration = value; break;
                case "meta:user-defined": if (name is not null) meta.UserDefined[name] = value; break;
            }
        }
        return meta;
    }

    /// <summary>An xsd:dateTime / xsd:date value (ODF 1.3 part 3 18.3.10 / 18.3.11); a zone suffix is dropped (the local wall time is shown).</summary>
    public static DateTime? ParseDate(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        string text = value!.Trim();
        if (text.EndsWith("Z", StringComparison.Ordinal)) text = text.Substring(0, text.Length - 1);
        int zone = text.LastIndexOfAny(['+', '-']);
        if (zone > 10) text = text.Substring(0, zone);       // past the date's own '-'
        int fraction = text.IndexOf('.');
        if (fraction > 0) text = text.Substring(0, fraction);
        string[] formats = ["yyyy-MM-dd'T'HH:mm:ss", "yyyy-MM-dd'T'HH:mm", "yyyy-MM-dd"];
        return DateTime.TryParseExact(text, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsed) ? parsed : null;
    }

    /// <summary>An xsd:duration time value such as PT01H02M03S (text:time-value of a fixed time field, 19.881).</summary>
    public static DateTime? ParseTime(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        string text = value!.Trim();
        if (text.StartsWith("PT", StringComparison.Ordinal))
        {
            try { TimeSpan span = XmlConvert.ToTimeSpan(text); return span < TimeSpan.Zero || span.TotalDays >= 1 ? null : DateTime.MinValue + span; }
            catch (Exception exception) when (exception is FormatException or OverflowException) { return null; }
        }
        if (ParseDate(text) is { } dateTime) return dateTime;
        return TimeSpan.TryParse(text, CultureInfo.InvariantCulture, out TimeSpan time) && time >= TimeSpan.Zero && time.TotalDays < 1 ? DateTime.MinValue + time : null;
    }
}

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
                case "meta:creation-date": meta.CreationDate = ParseDate(value); break;
                case "dc:date": meta.ModificationDate = ParseDate(value); break;
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

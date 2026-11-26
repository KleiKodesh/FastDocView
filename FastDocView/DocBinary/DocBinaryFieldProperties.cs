using System.Text;

namespace FastDocView.DocBinary;

/// <summary>
/// The document properties .doc fields show, from the compound file's property set streams ([MS-OLEPS]):
/// "\u0005SummaryInformation" (title PID 2, subject 3, author 4, keywords 5, comments 6, template 7, last author 8,
/// revision 9, edit time 10, last printed 11, created 12, last saved 13, words 15, characters 16) and
/// "\u0005DocumentSummaryInformation" (category 2, manager 14, company 15; its second section the user-defined properties,
/// named by the dictionary, PID 0) - the PIDs as Apache POI's PropertyIDMap.java:50-230 lists them; the stream layout as
/// POI's PropertySet.java:419-431 / Section.java:122-233 read it: a 28-byte header, then (FMTID, offset) per section; a
/// section: size, count, (PID, offset) pairs; a property: a 4-byte type (VT_I2 2, VT_I4 3, VT_LPSTR 30 in the section's
/// code page, VT_LPWSTR 31, VT_FILETIME 64 - Variant.java:51-202) and its value. Read once, on first use (small streams).
/// </summary>
internal sealed class DocBinaryFieldProperties
{
    private readonly Dictionary<string, string> _properties = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, DateTime> _dates = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// The properties dialog's values (<see cref="WordDocumentProperties"/>) from the same two streams: the summary set's
    /// PIDSI_* ([MS-OLEPS] 2.25.1 SummaryInformation: 0x0E page count, 0x0F word count, 0x10 character count, VT_I4; 0x12
    /// application name, VT_LPSTR) and the document summary set's PIDDSI_* ([MS-OSHARED] 2.3.3.2.1.1: 0x05 line count, 0x06
    /// paragraph count, 0x11 characters with spaces, VT_I4; 0x1B content status, a string). LibreOffice reads only the texts,
    /// dates and edit time of these sets (sfx2/source/doc/docinf.cxx LoadOlePropertySet), not the counts.
    /// </summary>
    public WordDocumentProperties Document { get; } = WordDocumentProperties.None;

    public DocBinaryFieldProperties(Func<string, byte[]?> readStream)
    {
        var summaryValues = new Dictionary<uint, object>();
        var documentSummaryValues = new Dictionary<uint, object>();
        if (readStream("\u0005SummaryInformation") is { } summary)
            foreach (var section in Sections(summary).Take(1))
                foreach ((uint id, object value) in section.properties)
                {
                    summaryValues[id] = value;
                    string? name = id switch
                    {
                        2 => "TITLE", 3 => "SUBJECT", 4 => "AUTHOR", 5 => "KEYWORDS", 6 => "COMMENTS", 7 => "TEMPLATE", 8 => "LASTSAVEDBY",
                        9 => "REVNUM", 15 => "NUMWORDS", 16 => "NUMCHARS", _ => null,
                    };
                    if (name is not null && value is string or int) _properties[name] = Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture)!;
                    // EDITTIME is a duration stored as a FILETIME (100 ns ticks): Word shows minutes
                    if (id == 10 && value is long ticks) _properties["EDITTIME"] = (ticks / TimeSpan.TicksPerMinute).ToString(System.Globalization.CultureInfo.InvariantCulture);
                    if (id is 11 or 12 or 13 && value is long fileTime && fileTime > 0)
                        _dates[id == 11 ? "PRINTDATE" : id == 12 ? "CREATEDATE" : "SAVEDATE"] = DateTime.FromFileTimeUtc(fileTime).ToLocalTime();
                }
        if (readStream("\u0005DocumentSummaryInformation") is { } documentSummary)
        {
            var sections = Sections(documentSummary).ToList();
            if (sections.Count > 0)
                foreach ((uint id, object value) in sections[0].properties)
                {
                    documentSummaryValues[id] = value;
                    string? name = id switch { 2 => "CATEGORY", 14 => "MANAGER", 15 => "COMPANY", _ => null };
                    if (name is not null && value is string text) _properties[name] = text;
                }
            if (sections.Count > 1 && sections[1].dictionary is { } names)
                foreach ((uint id, object value) in sections[1].properties)
                    if (names.TryGetValue(id, out string? name) && value is not null)
                        _properties[name] = value is long time ? DateTime.FromFileTimeUtc(time).ToLocalTime().ToString(System.Globalization.CultureInfo.InvariantCulture)
                                                               : Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture)!;
        }
        if (summaryValues.Count > 0 || documentSummaryValues.Count > 0)
            Document = DocumentProperties(summaryValues, documentSummaryValues);
    }

    private static WordDocumentProperties DocumentProperties(Dictionary<uint, object> summary, Dictionary<uint, object> documentSummary)
    {
        string? Text(Dictionary<uint, object> set, uint id) => set.TryGetValue(id, out object? value) && value is string text ? WordDocumentProperties.Text(text) : null;
        int? Count(Dictionary<uint, object> set, uint id) => set.TryGetValue(id, out object? value) && value is int count && count >= 0 ? count : null;
        DateTime? Moment(uint id) => summary.TryGetValue(id, out object? value) && value is long fileTime && fileTime > 0 ? DateTime.FromFileTimeUtc(fileTime) : null;
        // PIDSI_EDITTIME: a VT_FILETIME holding a duration in 100 ns ticks ([MS-OLEPS] 2.25.1); docinf.cxx subtracts the
        // 1601 epoch from the converted date - the same value
        int? editMinutes = summary.TryGetValue(10, out object? edit) && edit is long ticks && ticks > 0 && ticks / TimeSpan.TicksPerMinute <= int.MaxValue
            ? (int)(ticks / TimeSpan.TicksPerMinute) : null;
        return new WordDocumentProperties
        {
            Title = Text(summary, 2), Subject = Text(summary, 3), Author = Text(summary, 4), Keywords = Text(summary, 5),
            Comments = Text(summary, 6), Template = Text(summary, 7), LastModifiedBy = Text(summary, 8), Revision = Text(summary, 9),
            TotalEditingMinutes = editMinutes, LastPrinted = Moment(11), Created = Moment(12), Modified = Moment(13),
            Pages = Count(summary, 14), Words = Count(summary, 15), Characters = Count(summary, 16), Application = Text(summary, 18),
            Category = Text(documentSummary, 2), Lines = Count(documentSummary, 5), Paragraphs = Count(documentSummary, 6),
            Manager = Text(documentSummary, 14), Company = Text(documentSummary, 15), CharactersWithSpaces = Count(documentSummary, 0x11),
            ContentStatus = Text(documentSummary, 0x1B),
        };
    }

    public string? Property(string name) => _properties.TryGetValue(name, out string? value) ? value : null;
    public DateTime? Date(string name) => _dates.TryGetValue(name, out DateTime value) ? value : null;

    private static IEnumerable<(List<(uint id, object value)> properties, Dictionary<uint, string>? dictionary)> Sections(byte[] stream)
    {
        if (stream.Length < 28 || BitConverter.ToUInt16(stream, 0) != 0xFFFE) yield break;
        uint count = BitConverter.ToUInt32(stream, 24);
        for (int s = 0; s < count && 28 + s * 20 + 20 <= stream.Length; s++)
        {
            int offset = (int)BitConverter.ToUInt32(stream, 28 + s * 20 + 16);
            if (offset < 0 || offset + 8 > stream.Length) yield break;
            yield return ReadSection(stream, offset);
        }
    }

    private static (List<(uint, object)>, Dictionary<uint, string>?) ReadSection(byte[] stream, int offset)
    {
        var properties = new List<(uint, object)>();
        Dictionary<uint, string>? dictionary = null;
        int size = (int)Math.Min(BitConverter.ToUInt32(stream, offset), (uint)(stream.Length - offset));
        int propertyCount = (int)BitConverter.ToUInt32(stream, offset + 4);
        var entries = new List<(uint id, int at)>();
        for (int i = 0; i < propertyCount && offset + 8 + i * 8 + 8 <= offset + size; i++)
            entries.Add((BitConverter.ToUInt32(stream, offset + 8 + i * 8), offset + (int)BitConverter.ToUInt32(stream, offset + 12 + i * 8)));
        int codePage = 1252;
        foreach ((uint id, int at) in entries)
            if (id == 1 && at + 6 <= stream.Length) codePage = BitConverter.ToUInt16(stream, at + 4);       // PID_CODEPAGE, VT_I2
        Encoding encoding = EncodingOf(codePage);
        foreach ((uint id, int at) in entries)
        {
            if (at < 0 || at + 4 > stream.Length) continue;
            if (id == 0) { dictionary = ReadDictionary(stream, at, codePage, encoding); continue; }
            if (id == 1) continue;
            ushort type = BitConverter.ToUInt16(stream, at);
            int value = at + 4;
            try
            {
                switch (type)
                {
                    case 2: properties.Add((id, (int)BitConverter.ToInt16(stream, value))); break;
                    case 3: properties.Add((id, BitConverter.ToInt32(stream, value))); break;
                    case 11: properties.Add((id, BitConverter.ToInt16(stream, value) != 0 ? "Yes" : "No")); break;   // VT_BOOL
                    case 30:
                    {
                        int length = (int)BitConverter.ToUInt32(stream, value);
                        if (length < 0 || value + 4 + length > stream.Length) break;
                        properties.Add((id, (codePage == 1200 ? Encoding.Unicode.GetString(stream, value + 4, length) : encoding.GetString(stream, value + 4, length)).TrimEnd('\0')));
                        break;
                    }
                    case 31:
                    {
                        int characters = (int)BitConverter.ToUInt32(stream, value);
                        if (characters < 0 || value + 4 + characters * 2 > stream.Length) break;
                        properties.Add((id, Encoding.Unicode.GetString(stream, value + 4, characters * 2).TrimEnd('\0')));
                        break;
                    }
                    case 64: properties.Add((id, BitConverter.ToInt64(stream, value))); break;
                    case 5: properties.Add((id, BitConverter.ToDouble(stream, value).ToString(System.Globalization.CultureInfo.InvariantCulture))); break;   // VT_R8
                }
            }
            catch (ArgumentException) { }
        }
        return (properties, dictionary);
    }

    /// <summary>PID 0: count, then (PID, length, name) - in the code page's characters (UTF-16, length in characters, padded to 4 bytes, when 1200).</summary>
    private static Dictionary<uint, string> ReadDictionary(byte[] stream, int at, int codePage, Encoding encoding)
    {
        var names = new Dictionary<uint, string>();
        if (at + 4 > stream.Length) return names;
        uint count = BitConverter.ToUInt32(stream, at);
        int position = at + 4;
        for (int i = 0; i < count && position + 8 <= stream.Length; i++)
        {
            uint id = BitConverter.ToUInt32(stream, position);
            int length = (int)BitConverter.ToUInt32(stream, position + 4);
            position += 8;
            int bytes = codePage == 1200 ? length * 2 : length;
            if (length < 0 || position + bytes > stream.Length) break;
            names[id] = (codePage == 1200 ? Encoding.Unicode.GetString(stream, position, bytes) : encoding.GetString(stream, position, bytes)).TrimEnd('\0');
            position += bytes;
            if (codePage == 1200 && position % 4 != 0) position += 4 - position % 4;
        }
        return names;
    }

    private static Encoding EncodingOf(int codePage)
    {
        try { return codePage == 1200 ? Encoding.Unicode : Markup.MarkupTextEncoding.FromCodePage(codePage); }
        catch (Exception exception) when (exception is ArgumentException or NotSupportedException) { return Markup.MarkupTextEncoding.FromCodePage(1252); }
    }
}

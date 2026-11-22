using System.Text;

namespace FastDocView.Odt;

/// <summary>
/// The page setup of a .fodt without reading its head as XML: page layouts live in office:automatic-styles and master pages in
/// office:master-styles (ODF 1.3 part 3 3.15.3 / 3.15.4), so a byte search finds those two sections by the office prefix's tags
/// and only they are parsed, inside the root start tag (its namespace declarations). office:styles, which a LibreOffice
/// conversion of a Word file fills with 15 MB of list styles, is never tokenised. The search stops at office:body. Null when the
/// head does not have that shape (no root tag, no office prefix, a section left open): the caller reads the head whole.
/// </summary>
internal static class OdtFlatPageSetupReader
{
    private const int ChunkSize = 1 << 16;
    private const int MaxRootTagBytes = 1 << 20;
    private const string OfficeNamespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";

    public static OdtPageSetup? TryRead(Stream file)
    {
        file.Position = 0;
        var head = new MemoryStream();
        var chunk = new byte[ChunkSize];
        int rootEnd = -1;
        while (rootEnd < 0 && head.Length < MaxRootTagBytes)
        {
            int read = file.Read(chunk, 0, chunk.Length);
            if (read <= 0) return null;
            head.Write(chunk, 0, read);
            rootEnd = OdtStylesListStyleFilterStream.RootStartTagEnd(head.GetBuffer(), (int)head.Length);
        }
        if (rootEnd < 0) return null;
        string beforeRootEnd = Encoding.UTF8.GetString(head.GetBuffer(), 0, rootEnd);
        string rootTag = beforeRootEnd.Substring(beforeRootEnd.LastIndexOf('<'));
        string? office = OdtStylesListStyleFilterStream.PrefixOf(rootTag, OfficeNamespace);
        if (office is null) return null;
        string rootName = OdtXmlFragment.RootElementName(rootTag);

        byte[] marker = Encoding.UTF8.GetBytes(office + ":");
        var sections = new MemoryStream();
        long sectionStart = -1;                                  // file offset of the open section's '<' (-1: none open)
        string? openSection = null;
        // walk the file from the root tag's end; data = held-back tail + the next chunk
        long dataOffset = rootEnd;
        file.Position = rootEnd;
        byte[] data = [];
        int dataLength = 0;
        while (true)
        {
            int read = file.Read(chunk, 0, chunk.Length);
            if (read <= 0) return null;                          // no office:body: not the shape we read
            byte[] joined = new byte[dataLength + read];
            Buffer.BlockCopy(data, 0, joined, 0, dataLength);
            Buffer.BlockCopy(chunk, 0, joined, dataLength, read);
            data = joined;
            dataLength = joined.Length;
            int position = 0;
            int safeEnd = dataLength - marker.Length - 32;       // a tag name must fit after a match
            while (position < safeEnd)
            {
                int found = data.AsSpan(position, safeEnd - position).IndexOf(marker);
                if (found < 0) { position = safeEnd; break; }
                int at = position + found;
                position = at + marker.Length;
                bool isClose = at >= 2 && data[at - 1] == (byte)'/' && data[at - 2] == (byte)'<';
                bool isOpen = at >= 1 && data[at - 1] == (byte)'<';
                if (!isClose && !isOpen) continue;               // an attribute (office:version ...)
                string? local = LocalName(data, position, dataLength);
                if (local is null) continue;
                long tagStart = dataOffset + at - (isClose ? 2 : 1);
                if (isOpen && local == "body")
                {
                    if (openSection is not null) return null;
                    return Parse(rootTag, rootName, sections);
                }
                if (local is not ("automatic-styles" or "master-styles")) continue;
                if (isOpen && openSection is null)
                {
                    int tagEnd = Array.IndexOf(data, (byte)'>', position);
                    if (tagEnd > 0 && data[tagEnd - 1] == (byte)'/') continue;     // empty section
                    openSection = local;
                    sectionStart = tagStart;
                }
                else if (isClose && openSection == local)
                {
                    int tagEnd = Array.IndexOf(data, (byte)'>', position);
                    if (tagEnd < 0) return null;
                    long end = dataOffset + tagEnd + 1;
                    CopyRange(file, sectionStart, end, sections);
                    file.Position = dataOffset + dataLength;     // CopyRange moved it
                    openSection = null;
                }
            }
            // keep the unsearched tail for the next round
            // (two bytes before it too: the "<" or "</" of a marker that starts the tail)
            position = Math.Max(0, Math.Min(position, dataLength) - 2);
            int keep = dataLength - position;
            byte[] tail = new byte[keep];
            Buffer.BlockCopy(data, position, tail, 0, keep);
            dataOffset += position;
            data = tail;
            dataLength = keep;
        }
    }

    /// <summary>The element name after "prefix:" when it is followed by whitespace, '>' or '/'; else null.</summary>
    private static string? LocalName(byte[] data, int start, int length)
    {
        int end = start;
        while (end < length && (data[end] is >= (byte)'a' and <= (byte)'z' or (byte)'-')) end++;
        if (end == start || end >= length) return null;
        if (data[end] is not ((byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n' or (byte)'>' or (byte)'/')) return null;
        return Encoding.ASCII.GetString(data, start, end - start);
    }

    private static void CopyRange(Stream file, long start, long end, MemoryStream into)
    {
        file.Position = start;
        var buffer = new byte[ChunkSize];
        for (long left = end - start; left > 0;)
        {
            int read = file.Read(buffer, 0, (int)Math.Min(buffer.Length, left));
            if (read <= 0) break;
            into.Write(buffer, 0, read);
            left -= read;
        }
    }

    private static OdtPageSetup Parse(string rootTag, string rootName, MemoryStream sections)
    {
        var wrapped = new MemoryStream();
        byte[] open = Encoding.UTF8.GetBytes(rootTag), close = Encoding.UTF8.GetBytes("</" + rootName + ">");
        wrapped.Write(open, 0, open.Length);
        sections.Position = 0;
        sections.CopyTo(wrapped);
        wrapped.Write(close, 0, close.Length);
        wrapped.Position = 0;
        return OdtPageSetup.Read(wrapped);
    }
}

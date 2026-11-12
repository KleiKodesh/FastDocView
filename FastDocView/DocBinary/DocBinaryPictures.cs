using FastDocView.Html;

namespace FastDocView.DocBinary;

/// <summary>
/// Inline pictures of a .doc: a 0x01 character with sprmCFSpec whose sprmCPicLocation (0x6A03) names a PICF in the Data
/// stream, as LibreOffice reads them (ww8par.cxx "case 0x1": a picture when fSpec is set and the run is not an OLE object;
/// ww8par6.cxx Read_PicLoc; ww8graf2.cxx ImportGraf / PicRead). Only the picture at hand is read - its PICF header and the
/// OfficeArt record headers after it, then (for <see cref="Read"/>) its blip bytes; nothing is kept.
///
/// PICF (ww8struc.hxx WW8_PIC_SHADOW, 0x2E bytes, then 4 BRCs, dxaOrigin, dyaOrigin, cProps - PicRead): lcb at 0x0 (the
/// PICF and its data), cbHeader at 0x4, mfp.mm at 0x6, dxaGoal / dyaGoal at 0x1C / 0x1E (twips), mx / my at 0x20 / 0x22
/// (scale in 0.1 %), crops at 0x24..0x2A. Shown size (ww8graf2.cxx WW8PicDesc): (goal - crops) * scale / 1000.
/// mm 0x64 / 0x66 = an OfficeArt picture (0x66: a name of one length byte first): an OfficeArtSpContainer, then the FBSE
/// that holds the blip (msdffimp.cxx: "look at the end of this record for a FBSE pool": the blip starts 36 bytes into the
/// FBSE record's body). Blip records (msdffimp.cxx GetBLIPDirect): recInstance & 0xFFFE = PNG 0x6E0, JPEG 0x46A / 0x6E2,
/// DIB 0x7A8, TIFF 0x6E4, EMF 0x3D4, WMF 0x216, PICT 0x542; recInstance &amp; 1 = a second 16-byte UID; bitmaps have one
/// tag byte before the data. Record header (msdffimp.cxx ReadCommonRecordHeader): u16 ver (low 4 bits) / instance (high
/// 12), u16 type, u32 length. Only PNG and JPEG are shown (a DIB has no file header, metafiles are compressed vector
/// formats a browser cannot draw).
/// </summary>
internal sealed class DocBinaryPictures
{
    private const int PicfMinimumLength = 58;     // ww8graf2.cxx ImportGraf: "aPic.lcb >= 58" sanity check
    private const int FbseBlipOffset = 36;        // msdffimp.cxx: 20 + 4 + 4 + 4 + 4 bytes into the FBSE body
    private readonly Stream? _dataStream;

    public DocBinaryPictures(Stream? dataStream) => _dataStream = dataStream;

    /// <summary>
    /// A picture's shown size (points), its crop and where its blip's data lies in the Data stream; null when it is not a shown
    /// PNG / JPEG. The crop: the picture's OfficeArt cropFrom* properties when its shape has them - where Word 16 keeps it
    /// (test-data/word-picture-crop: PICF crops 0, the size by mx / my) and LibreOffice reads it for an OfficeArt picture
    /// (ww8graf2.cxx:524-685 ImportObj, then SetAttributesAtGrfNode, ww8graf.cxx:2180-2222) - else the PICF's dxaCropLeft /
    /// dyaCropTop / dxaCropRight / dyaCropBottom (twips over the goal size, negative = padding; ww8graf2.cxx:303-318
    /// WW8PicDesc, :431 SwCropGrf).
    /// </summary>
    public (double widthPt, double heightPt, long dataAt, int dataLength, string extension, WordPictureCrop crop)? Locate(uint pictureLocation)
    {
        Stream? data = _dataStream;
        if (data is null || pictureLocation + 0x44L > data.Length) return null;
        byte[] header = Read(data, pictureLocation, 0x44);
        if (header.Length < 0x44) return null;
        uint lcb = BitConverter.ToUInt32(header, 0);
        int cbHeader = BitConverter.ToUInt16(header, 4);
        int mm = BitConverter.ToUInt16(header, 6);
        if (lcb < PicfMinimumLength || cbHeader < 0x44 || cbHeader > lcb || pictureLocation + (long)lcb > data.Length) return null;
        if (mm is not (0x64 or 0x66)) return null;
        double Twips(int at) => BitConverter.ToInt16(header, at);
        double goalWidth = Twips(0x1C) - Twips(0x24) - Twips(0x28), goalHeight = Twips(0x1E) - Twips(0x26) - Twips(0x2A);
        double widthPt = Math.Max(1, goalWidth) * BitConverter.ToUInt16(header, 0x20) / 1000 / 20;
        double heightPt = Math.Max(1, goalHeight) * BitConverter.ToUInt16(header, 0x22) / 1000 / 20;
        WordPictureCrop crop = WordPictureCrop.FromLengths(Twips(0x24), Twips(0x26), Twips(0x28), Twips(0x2A), Twips(0x1C), Twips(0x1E));

        long position = pictureLocation + cbHeader, end = pictureLocation + (long)lcb;
        if (mm == 0x66)
        {
            byte[] nameLength = Read(data, position, 1);
            if (nameLength.Length < 1) return null;
            position += 1 + nameLength[0];
        }
        // the OfficeArtSpContainer, then the records after it: the first FBSE holds the blip
        if (RecordHeader(data, position) is not { Type: 0xF004 } shape) return null;
        if (ShapeCrop(data, position + 8, Math.Min(end, position + 8 + shape.Length)) is { } shapeCrop) crop = shapeCrop;
        position += 8 + shape.Length;
        while (position + 8 <= end)
        {
            if (RecordHeader(data, position) is not { } record) return null;
            if (record.Type == 0xF007 && record.Length >= FbseBlipOffset + 8)
            {
                if (Blip(data, position + 8 + FbseBlipOffset) is not { } blip) return null;
                return (widthPt, heightPt, blip.dataAt, blip.dataLength, blip.extension, crop);
            }
            position += 8 + record.Length;
        }
        return null;
    }

    /// <summary>
    /// The Office Math (OMML) a picture stands for: Word saves an equation in a .doc as a picture whose OfficeArt shape
    /// carries the equationXML property (MS-ODRAW 2.3.2.10 / 2.3.2.11: opid 0x030C, complex, a UTF-8 WordprocessingML 2003
    /// document holding the m:oMath / m:oMathPara). Its property tables (FOPT 0xF00B, secondary 0xF121, tertiary 0xF122:
    /// 6-byte entries, then the complex data in entry order, MS-ODRAW 2.2.9 / 2.3.1) are read from the picture's
    /// OfficeArtSpContainer. LibreOffice does not read this property (its .doc import shows the picture); Word does. Null: none.
    /// </summary>
    public string? EquationXml(uint pictureLocation)
    {
        Stream? data = _dataStream;
        if (data is null || pictureLocation + 0x44L > data.Length) return null;
        byte[] header = Read(data, pictureLocation, 0x44);
        if (header.Length < 0x44) return null;
        uint lcb = BitConverter.ToUInt32(header, 0);
        int cbHeader = BitConverter.ToUInt16(header, 4);
        int mm = BitConverter.ToUInt16(header, 6);
        if (lcb < PicfMinimumLength || cbHeader < 0x44 || cbHeader > lcb || pictureLocation + (long)lcb > data.Length || mm is not (0x64 or 0x66)) return null;
        long position = pictureLocation + cbHeader;
        if (mm == 0x66)
        {
            byte[] nameLength = Read(data, position, 1);
            if (nameLength.Length < 1) return null;
            position += 1 + nameLength[0];
        }
        if (RecordHeader(data, position) is not { Type: 0xF004 } shape) return null;
        long end = Math.Min(position + 8 + shape.Length, pictureLocation + (long)lcb);
        for (position += 8; position + 8 <= end;)
        {
            if (RecordHeader(data, position) is not { } record) return null;
            if (record.Type is 0xF00B or 0xF121 or 0xF122 && record.Length <= end - position - 8)
            {
                byte[] table = Read(data, position + 8, (int)record.Length);
                int entries = Math.Min(record.Instance, table.Length / 6);
                long complexAt = entries * 6L;
                for (int i = 0; i < entries; i++)
                {
                    int id = BitConverter.ToUInt16(table, i * 6);
                    uint value = BitConverter.ToUInt32(table, i * 6 + 2);
                    bool isComplex = (id & 0x8000) != 0;
                    if ((id & 0x3FFF) == 0x030C && isComplex && complexAt + value <= table.Length)
                        return System.Text.Encoding.UTF8.GetString(table, (int)complexAt, (int)value);
                    if (isComplex) complexAt += value;
                }
            }
            position += 8 + record.Length;
        }
        return null;
    }

    /// <summary>
    /// The cropFromTop / Bottom / Left / Right properties (msdffdef.hxx:131-134, ids 256-259: signed 16.16 fractions of the
    /// picture) of the FOPT in the SpContainer body [<paramref name="start"/>, <paramref name="end"/>); null when it sets none.
    /// A whole part of 50 or more is dropped as LibreOffice drops it (ww8graf.cxx:2169-2174, values old OOo mangled).
    /// </summary>
    internal static WordPictureCrop? ShapeCrop(Stream data, long start, long end)
    {
        for (long position = start; position + 8 <= end;)
        {
            if (RecordHeader(data, position) is not { } record) return null;
            if (record.Type == 0xF00B)
            {
                byte[] table = Read(data, position + 8, (int)Math.Min(record.Length, (long)record.Instance * 6));
                double[] crop = new double[4];   // top, bottom, left, right
                bool isSet = false;
                for (int i = 0; i + 6 <= table.Length; i += 6)
                {
                    int property = BitConverter.ToUInt16(table, i) & 0x3FFF;
                    if (property is < 256 or > 259) continue;
                    int value = BitConverter.ToInt32(table, i + 2);
                    crop[property - 256] = Math.Abs(value >> 16) >= 50 ? 0 : value / 65536.0;
                    isSet = true;
                }
                return isSet ? new WordPictureCrop(crop[2], crop[0], crop[3], crop[1]) : null;
            }
            position += 8 + record.Length;
        }
        return null;
    }

    /// <summary>The PNG / JPEG data of the blip record at <paramref name="blipAt"/> in <paramref name="data"/>; null for
    /// another record, another blip type or data past the stream's end.</summary>
    internal static (long dataAt, int dataLength, string extension)? Blip(Stream data, long blipAt)
    {
        if (RecordHeader(data, blipAt) is not { } blip || blip.Type is < 0xF018 or > 0xF117) return null;
        string? extension = (blip.Instance & 0xFFFE) switch { 0x6E0 => ".png", 0x46A or 0x6E2 => ".jpg", _ => null };
        if (extension is null) return null;
        int skip = ((blip.Instance & 1) != 0 ? 32 : 16) + 1;
        long dataAt = blipAt + 8 + skip;
        long dataLength = blip.Length - skip;
        if (dataLength <= 0 || dataAt + dataLength > data.Length || dataLength > int.MaxValue) return null;
        return (dataAt, (int)dataLength, extension);
    }

    /// <summary>A DIB blip's data (recInstance &amp; 0xFFFE = 0x7A8; the BITMAPINFOHEADER and pixels after the UID(s) and the tag byte - msdffimp.cxx GetBLIPDirect).</summary>
    internal static (long dataAt, int dataLength)? DibBlip(Stream data, long blipAt)
    {
        if (RecordHeader(data, blipAt) is not { } blip || blip.Type is < 0xF018 or > 0xF117 || (blip.Instance & 0xFFFE) != 0x7A8) return null;
        int skip = ((blip.Instance & 1) != 0 ? 32 : 16) + 1;
        long dataAt = blipAt + 8 + skip, dataLength = blip.Length - skip;
        if (dataLength <= 0 || dataAt + dataLength > data.Length || dataLength > int.MaxValue) return null;
        return (dataAt, (int)dataLength);
    }

    /// <summary>The picture's bytes, or null (see <see cref="Locate"/>).</summary>
    public WordDocumentPicture? Read(uint pictureLocation)
    {
        if (Locate(pictureLocation) is not { } picture || _dataStream is null) return null;
        byte[] bytes = Read(_dataStream, picture.dataAt, picture.dataLength);
        return bytes.Length == picture.dataLength ? WordPictureHtml.Picture(picture.extension, bytes) : null;
    }

    internal readonly record struct OfficeArtRecordHeader(int Instance, int Type, long Length);

    internal static OfficeArtRecordHeader? RecordHeader(Stream data, long at)
    {
        byte[] header = Read(data, at, 8);
        if (header.Length < 8) return null;
        int versionAndInstance = BitConverter.ToUInt16(header, 0);
        return new OfficeArtRecordHeader(versionAndInstance >> 4, BitConverter.ToUInt16(header, 2), BitConverter.ToUInt32(header, 4));
    }

    internal static byte[] Read(Stream data, long at, int length)
    {
        if (at < 0 || at >= data.Length) return [];
        data.Position = at;
        var bytes = new byte[(int)Math.Min(length, data.Length - at)];
        int read = 0;
        while (read < bytes.Length)
        {
            int got = data.Read(bytes, read, bytes.Length - read);
            if (got <= 0) break;
            read += got;
        }
        return read == bytes.Length ? bytes : bytes.AsSpan(0, read).ToArray();
    }
}

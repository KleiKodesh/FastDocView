using System.Text;

namespace FastDocView.DocBinary;

/// <summary>
/// SttbfFfn ([MS-DOC] §2.9.284): the document's font names by index (ftc), which the character
/// sprms sprmCRgFtc0 / sprmCRgFtc1 / sprmCRgFtc2 / sprmCFtcBi point at. Each entry is an FFN
/// (§2.9.94): a length byte, 39 bytes of metrics and flags, then the name as null-terminated UTF-16.
/// </summary>
public sealed class DocBinaryFontTable
{
    private const int NameOffsetInEntry = 1 + 39;

    private readonly string[] _names;

    public static DocBinaryFontTable Empty { get; } = new([]);

    private DocBinaryFontTable(string[] names) { _names = names; }

    public string? NameOf(int fontIndex) => fontIndex >= 0 && fontIndex < _names.Length && _names[fontIndex].Length > 0 ? _names[fontIndex] : null;

    public int Count => _names.Length;

    public static DocBinaryFontTable Read(Stream tableStream, (uint offset, uint length) location)
    {
        if (location.length < 4 || DocBinaryStructureReader.TryRead(tableStream, location) is not { } bytes) return Empty;   // none, or outside the table stream

        int count = BitConverter.ToUInt16(bytes, 0);
        int position = 4;                                     // cData, cbExtra
        var names = new string[count];
        for (int i = 0; i < count && position < bytes.Length; i++)
        {
            int entryLength = bytes[position] + 1;            // cbFfnM1 counts the bytes after itself
            int nameStart = position + NameOffsetInEntry;
            int entryEnd = Math.Min(bytes.Length, position + entryLength);
            var name = new StringBuilder();
            for (int at = nameStart; at + 1 < entryEnd; at += 2)
            {
                char c = (char)BitConverter.ToUInt16(bytes, at);
                if (c == '\0') break;
                name.Append(c);
            }
            names[i] = name.ToString();
            position += entryLength;
        }
        return new DocBinaryFontTable(names);
    }
}

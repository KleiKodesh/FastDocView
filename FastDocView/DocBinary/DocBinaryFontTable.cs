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

    private DocBinaryFontTable(string[] names, byte[]? characterSets = null) { _names = names; CharacterSets = characterSets ?? []; }

    /// <summary>Each font's character set (FFN chs); filled for Word 6/95 files, where it picks the code page of 8-bit text.</summary>
    public IReadOnlyList<byte> CharacterSets { get; }

    public string? NameOf(int fontIndex) => fontIndex >= 0 && fontIndex < _names.Length && _names[fontIndex].Length > 0 ? _names[fontIndex] : null;

    public int Count => _names.Length;

    /// <summary>The first font index with this name; -1 when none.</summary>
    public int IndexOf(string name) => Array.IndexOf(_names, name);

    public static DocBinaryFontTable Read(Stream tableStream, (uint offset, uint length) location, bool isWord6 = false)
    {
        if (location.length < 4 || DocBinaryStructureReader.TryRead(tableStream, location) is not { } bytes) return Empty;   // none, or outside the table stream
        if (isWord6) return ReadWord6(bytes);

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

    /// <summary>
    /// Word 6/95 SttbfFfn (ww8scan.cxx WW8Fonts ctor, eVersion &lt; eWW8): the table's byte length (2 bytes), then FFNs until the
    /// end - no count. FFN: cbFfnM1, flags, wWeight (2), chs, ibszAlt, then the name as null-terminated 8-bit text in the
    /// font's own code page (#i8726#), cp1252 for an ANSI or symbol font. LibreOffice caps a name at 65 bytes (maxStrSize).
    /// </summary>
    private static DocBinaryFontTable ReadWord6(byte[] bytes)
    {
        var names = new List<string>();
        var characterSets = new List<byte>();
        int position = 2;
        while (position + 6 < bytes.Length)
        {
            int entryLength = bytes[position] + 1;
            int entryEnd = Math.Min(bytes.Length, position + entryLength);
            byte characterSet = bytes[position + 4];
            char[]? upperHalf = DocBinaryTextDecoder.UpperHalfOf(DocBinaryTextDecoder.CodePageOfCharacterSet(characterSet));
            var name = new StringBuilder();
            for (int at = position + 6; at < entryEnd && name.Length < 65 && bytes[at] != 0; at++)
                name.Append(bytes[at] < 0x80 ? (char)bytes[at] : upperHalf is null ? DocBinaryTextDecoder.DecodeCompressed(bytes[at]) : upperHalf[bytes[at] - 0x80]);
            names.Add(name.ToString());
            characterSets.Add(characterSet);
            position += entryLength;
        }
        return new DocBinaryFontTable(names.ToArray(), characterSets.ToArray());
    }
}

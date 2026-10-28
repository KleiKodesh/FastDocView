using System.Text;

namespace FastDocView.DocBinary;

/// <summary>
/// Decodes .doc piece text: UTF-16LE pieces as-is, "compressed" pieces as cp1252 with the 0x80–0x9F
/// mappings [MS-DOC] §2.4.1 lists. Hand-rolled so no code-page provider is needed on .NET 10.
/// </summary>
public static class DocBinaryTextDecoder
{
    private static readonly char[] HighCp1252 =
    [
        '€', '\u0081', '‚', 'ƒ', '„', '…', '†', '‡', 'ˆ', '‰', 'Š', '‹', 'Œ', '\u008D', 'Ž', '\u008F',
        '\u0090', '‘', '’', '“', '”', '•', '–', '—', '˜', '™', 'š', '›', 'œ', '\u009D', 'ž', 'Ÿ',
    ];

    public static char DecodeCompressed(byte value)
        => value is >= 0x80 and <= 0x9F ? HighCp1252[value - 0x80] : (char)value;

    public static int Decode(byte[] bytes, int byteCount, bool isCompressed, char[] destination)
    {
        if (isCompressed)
        {
            for (int i = 0; i < byteCount; i++) destination[i] = DecodeCompressed(bytes[i]);
            return byteCount;
        }
        return Encoding.Unicode.GetChars(bytes, 0, byteCount & ~1, destination, 0);
    }

    public static void Append(StringBuilder text, byte[] bytes, int byteCount, bool isCompressed)
    {
        if (isCompressed) for (int i = 0; i < byteCount; i++) text.Append(DecodeCompressed(bytes[i]));
        else text.Append(Encoding.Unicode.GetString(bytes, 0, byteCount & ~1));
    }
}

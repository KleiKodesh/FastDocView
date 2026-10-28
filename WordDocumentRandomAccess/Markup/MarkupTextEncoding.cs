using System.Text;

namespace WordDocumentRandomAccess.Markup;

/// <summary>
/// The character encodings the markup readers (.txt, MHTML) work with, by Windows code page, and the
/// .txt encoding detection. Detection follows Notepad++ (the user's reference for .txt): its order and
/// rules were read in its source (PowerEditor/src/ScintillaComponent/Buffer.cpp FileManager::loadFileData /
/// detectCodepage, PowerEditor/src/Utf8_16.cpp Utf8_16_Read::determineEncoding / utf8_7bits_8bits) and
/// reimplemented here, no code copied:
///   1. BOM: FE FF → UTF-16 BE, FF FE → UTF-16 LE, EF BB BF → UTF-8 (Notepad++ has no UTF-32 BOM check).
///   2. (Notepad++ runs uchardet here when "autodetect character encoding" is on — not reimplemented; see NOTES.md.)
///   3. UTF-16 LE without BOM: even length, byte0 != 0 and byte1 == 0, plus a statistics test
///      (Notepad++ calls Win32 IsTextUnicode; approximated: most odd bytes of the sample are 0).
///   4. Structural UTF-8 check on the first block (128 KB): lead bytes with the right continuation
///      count, any NUL = not UTF-8; 7-bit ASCII is UTF-8-compatible.
///   5. Otherwise the ANSI code page — windows-1255 for this Hebrew collection.
/// </summary>
public static class MarkupTextEncoding
{
    public const int Utf8 = 65001, Utf16LittleEndian = 1200, Utf16BigEndian = 1201, HebrewWindows = 1255;

    /// <summary>Notepad++ reads (and detects on) its first block of 128 KB + 4 bytes.</summary>
    public const int DetectionSampleBytes = 128 * 1024 + 4;

    public static Encoding FromCodePage(int codePage) => codePage switch
    {
        Utf8 => new UTF8Encoding(false, false),
        Utf16LittleEndian => new UnicodeEncoding(false, false, false),
        Utf16BigEndian => new UnicodeEncoding(true, false, false),
#if NETFRAMEWORK
        _ => Encoding.GetEncoding(codePage),
#else
        _ => CodePagesEncodingProvider.Instance.GetEncoding(codePage) ?? Encoding.GetEncoding(codePage),
#endif
    };

    /// <summary>A MIME / HTML charset label → code page; null when unknown.</summary>
    public static int? CodePageOfCharsetName(string name)
    {
        name = name.Trim().Trim('"', '\'').ToLowerInvariant();
        if (name.Length == 0) return null;
        if (name is "utf-8" or "utf8") return Utf8;
        if (name is "utf-16" or "utf-16le") return Utf16LittleEndian;
        if (name == "utf-16be") return Utf16BigEndian;
        if (name is "iso-8859-8-i" or "iso-8859-8") return 28598;     // same letters as windows-1255 for the base alphabet
        try
        {
#if NETFRAMEWORK
            return Encoding.GetEncoding(name).CodePage;
#else
            Encoding? encoding = CodePagesEncodingProvider.Instance.GetEncoding(name);
            return (encoding ?? Encoding.GetEncoding(name)).CodePage;
#endif
        }
        catch (ArgumentException) { return null; }
    }

    /// <summary>The encoding of a .txt from its first bytes (see the class summary). Returns the code page and the BOM length to skip.</summary>
    public static (int codePage, int bomLength, string reason) DetectPlainText(ReadOnlySpan<byte> sample, long fileLength)
    {
        if (sample.Length >= 2 && sample[0] == 0xFE && sample[1] == 0xFF) return (Utf16BigEndian, 2, "BOM FE FF");
        if (sample.Length >= 2 && sample[0] == 0xFF && sample[1] == 0xFE) return (Utf16LittleEndian, 2, "BOM FF FE");
        if (sample.Length >= 3 && sample[0] == 0xEF && sample[1] == 0xBB && sample[2] == 0xBF) return (Utf8, 3, "BOM EF BB BF");
        if (sample.Length == 0) return (Utf8, 0, "empty");
        if (LooksLikeUtf16LittleEndianWithoutBom(sample, fileLength)) return (Utf16LittleEndian, 0, "UTF-16 LE without BOM");
        return Utf8Structure(sample) switch
        {
            Utf8Kind.SevenBit => (Utf8, 0, "7-bit ASCII"),
            Utf8Kind.Utf8 => (Utf8, 0, "valid UTF-8"),
            _ => (HebrewWindows, 0, "8-bit, not UTF-8 → ANSI (windows-1255)"),
        };
    }

    private static bool LooksLikeUtf16LittleEndianWithoutBom(ReadOnlySpan<byte> sample, long fileLength)
    {
        if ((fileLength & 1) != 0 || sample.Length < 2 || sample[0] == 0 || sample[1] != 0) return false;
        int pairs = sample.Length / 2, zeroHigh = 0, zeroLow = 0;
        for (int i = 0; i + 1 < sample.Length; i += 2)
        {
            if (sample[i + 1] == 0) zeroHigh++;
            if (sample[i] == 0) zeroLow++;
        }
        return zeroHigh * 2 > pairs && zeroLow * 10 < pairs;
    }

    public enum Utf8Kind { SevenBit, Utf8, EightBit }

    /// <summary>Structural UTF-8 test as Notepad++ does it: right number of continuation bytes, no NUL; a sequence cut by the sample end is not held against it.</summary>
    public static Utf8Kind Utf8Structure(ReadOnlySpan<byte> sample)
    {
        bool sawHighByte = false;
        int i = 0;
        while (i < sample.Length)
        {
            byte b = sample[i];
            if (b == 0) return Utf8Kind.EightBit;
            if (b < 0x80) { i++; continue; }
            sawHighByte = true;
            int continuation = b >= 0xC0 && b <= 0xDF ? 1 : b >= 0xE0 && b <= 0xEF ? 2 : b >= 0xF0 && b <= 0xF7 ? 3 : -1;
            if (continuation < 0) return Utf8Kind.EightBit;
            for (int k = 1; k <= continuation; k++)
            {
                if (i + k >= sample.Length) return Utf8Kind.Utf8;
                if ((sample[i + k] & 0xC0) != 0x80) return Utf8Kind.EightBit;
            }
            i += continuation + 1;
        }
        return sawHighByte ? Utf8Kind.Utf8 : Utf8Kind.SevenBit;
    }

    public static bool ContainsHebrew(string text)
    {
        foreach (char c in text) if (c >= '֐' && c <= '׿') return true;
        return false;
    }
}

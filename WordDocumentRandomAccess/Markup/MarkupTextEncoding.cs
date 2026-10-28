using System.Text;
using TextEncodingDetection;

namespace WordDocumentRandomAccess.Markup;

/// <summary>
/// The character encodings the markup readers (.txt, MHTML) work with, by Windows code page, and the
/// .txt encoding detection. Detection is the TextEncodingDetection library (sibling repo) in its
/// default Improved mode: Notepad++'s decision (the user's reference for .txt) wherever Notepad++
/// is right, plus what it misses (UTF-32 BOMs, BOM-less UTF-16 BE/LE, every uchardet charset, ICU's
/// detector as a second opinion for non-Hebrew 8-bit text); windows-1255 stays the ANSI fallback for
/// this Hebrew collection. See its README.
/// </summary>
public static class MarkupTextEncoding
{
    public const int Utf8 = 65001, Utf16LittleEndian = 1200, Utf16BigEndian = 1201, HebrewWindows = 1255;

    /// <summary>Notepad++ reads (and detects on) its first block of 128 KB + 4 bytes.</summary>
    public const int DetectionSampleBytes = TextEncodingDetector.SampleSize;

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

    /// <summary>The encoding of a .txt from its first bytes (at least min(fileLength, <see cref="DetectionSampleBytes"/>)).
    /// Returns the code page, the BOM length to skip and the reason (stored in the index).</summary>
    public static (int codePage, int bomLength, string reason) DetectPlainText(ReadOnlySpan<byte> sample, long fileLength)
    {
        DetectionResult r = TextEncodingDetector.Detect(sample, fileLength, HebrewWindows);
        // The block scanner reads 1-byte, double-byte (lead/trail) and UTF-16 code units. For what it
        // cannot split safely (UTF-32, GB18030 four-byte, EUC-JP three-byte, stateful ISO-2022 / HZ)
        // the Notepad++ decision stands, as before Improved mode.
        if (!IsScannable(r.CodePage))
            r = TextEncodingDetector.Detect(sample, fileLength, HebrewWindows, DetectionMode.NotepadPlusPlus);
        string reason = r.Reason == DetectionReason.Uchardet ? $"uchardet {r.UchardetCharset} ({r.Confidence:0.00})" : r.Reason.ToString();
        return (r.CodePage, r.BomLength, reason);
    }

    /// <summary>Code pages <see cref="MarkupBlockScanner"/> handles: UTF-8, UTF-16 LE/BE, the
    /// double-byte ones it knows the lead bytes of, and every single-byte code page.</summary>
    public static bool IsScannable(int codePage) =>
        codePage is Utf8 or Utf16LittleEndian or Utf16BigEndian or 932 or 936 or 949 or 950 or 51949
        || (codePage is not (12000 or 12001 or 54936 or 51932 or 50220 or 50221 or 50222 or 50225 or 50227 or 52936)
            && FromCodePage(codePage).IsSingleByte);

    public static bool ContainsHebrew(string text)
    {
        foreach (char c in text) if (c >= '֐' && c <= '׿') return true;
        return false;
    }
}

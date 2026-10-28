using System.Text;
using TextEncodingDetection;

namespace WordDocumentRandomAccess.Markup;

/// <summary>
/// The character encodings the markup readers (.txt, MHTML) work with, by Windows code page, and the
/// .txt encoding detection. Detection is the TextEncodingDetection library (sibling repo), which
/// reproduces Notepad++ (the user's reference for .txt) exactly: BOM, then the uchardet copy
/// Notepad++ bundles (ported), then UTF-16 LE without BOM (IsTextUnicode statistics), UTF-8
/// structure and the ANSI code page — windows-1255 for this Hebrew collection. See its README.
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
        string reason = r.Reason == DetectionReason.Uchardet ? $"uchardet {r.UchardetCharset} ({r.Confidence:0.00})" : r.Reason.ToString();
        return (r.CodePage, r.BomLength, reason);
    }

    public static bool ContainsHebrew(string text)
    {
        foreach (char c in text) if (c >= '֐' && c <= '׿') return true;
        return false;
    }
}

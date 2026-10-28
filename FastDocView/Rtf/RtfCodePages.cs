using System.Text;

namespace FastDocView.Rtf;

/// <summary>
/// Code pages for \'xx bytes and 8-bit text. \fcharsetN → code page follows LibreOffice
/// rtfcharsets.cxx (which cites RTF spec 1.9.1 p.19): 0 → 1252, 1 → the document's \ansicpg, 2 (Symbol) → bytes
/// kept as U+F0xx, 177 → 1255 (Hebrew), 178 → 1256, ... Font-name suffixes "(Hebrew)", "Cyr", ... pick a code page
/// as LibreOffice's handleFontTableEntry does.
/// </summary>
internal static class RtfCodePages
{
    public const ushort Symbol = 42;
    public const ushort Windows1252 = 1252;

    private static readonly (int charset, int codePage)[] CharsetCodePages =
    [
        (0, 1252), (1, 0), (2, 42), (77, 10000), (78, 10001), (79, 10003), (80, 10008), (81, 10002), (83, 10005), (84, 10004), (85, 10006),
        (86, 10081), (87, 10021), (88, 10029), (89, 10007), (128, 932), (129, 949), (130, 1361), (134, 936), (136, 950), (161, 1253), (162, 1254),
        (163, 1258), (177, 1255), (178, 1256), (186, 1257), (204, 1251), (222, 874), (238, 1250), (254, 437), (255, 850),
    ];

    private static readonly (string suffix, int codePage)[] FontNameSuffixes =
    [
        ("Baltic", 1257), ("CE", 1250), ("Cyr", 1251), ("Greek", 1253), ("Tur", 1254), ("(Hebrew)", 1255), ("(Arabic)", 1256), ("(Vietnamese)", 1258),
    ];

    private static readonly Dictionary<int, Encoding?> Encodings = new();

    static RtfCodePages()
    {
#if NET
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
#endif
    }

    /// <summary>The code page of \fcharsetN; 0 = "the document default" (charset 1), -1 = unknown charset.</summary>
    public static int CodePageOfCharset(int charset)
    {
        foreach ((int c, int codePage) in CharsetCodePages) if (c == charset) return codePage;
        return -1;
    }

    /// <summary>"Arial (Hebrew)" → 1255 and the name without the suffix; LibreOffice handleFontTableEntry.</summary>
    public static (string name, int codePage) SplitFontNameSuffix(string name)
    {
        int lastSpace = name.LastIndexOf(' ');
        if (lastSpace < 0) return (name, -1);
        string suffix = name.Substring(lastSpace + 1);
        foreach ((string s, int codePage) in FontNameSuffixes)
            if (s == suffix) return (name.Substring(0, lastSpace), codePage);
        return (name, -1);
    }

    public static Encoding? EncodingOf(int codePage)
    {
        lock (Encodings)
        {
            if (Encodings.TryGetValue(codePage, out Encoding? cached)) return cached;
            Encoding? encoding;
            try { encoding = codePage is Symbol or <= 0 ? null : Encoding.GetEncoding(codePage); }
            catch (Exception exception) when (exception is ArgumentException or NotSupportedException) { encoding = null; }
            Encodings[codePage] = encoding;
            return encoding;
        }
    }

    /// <summary>Decodes 8-bit text bytes; the Symbol charset and unknown code pages map bytes to U+F0xx / Latin-1.</summary>
    public static int Decode(ReadOnlySpan<byte> bytes, int codePage, Span<char> destination)
    {
        if (codePage == Symbol)
        {
            for (int i = 0; i < bytes.Length; i++) destination[i] = (char)(0xF000 + bytes[i]);
            return bytes.Length;
        }
        Encoding? encoding = EncodingOf(codePage) ?? EncodingOf(Windows1252);
        if (encoding is null)
        {
            for (int i = 0; i < bytes.Length; i++) destination[i] = (char)bytes[i];
            return bytes.Length;
        }
#if NET
        return encoding.GetChars(bytes, destination);
#else
        byte[] array = bytes.ToArray();
        char[] chars = encoding.GetChars(array);
        chars.AsSpan().CopyTo(destination);
        return chars.Length;
#endif
    }
}

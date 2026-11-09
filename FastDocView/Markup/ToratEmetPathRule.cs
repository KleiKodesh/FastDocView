using System.Text;

namespace FastDocView.Markup;

/// <summary>
/// The path rule: a .txt whose full path holds the folder word "ToratEmet" (any folder level, any case) is a Torat Emet book with NO
/// detection - parsed the same way as a detected one, even when line 1 is not an "&amp;Key=" line and no line has a heading mark.
/// A separate entry point on purpose: <see cref="ToratEmetFormat.TryReadHead"/> (the detection) is untouched.
/// </summary>
public static class ToratEmetPathRule
{
    public const string FolderWord = "ToratEmet";

    public static bool Applies(string? filePath) =>
        !string.IsNullOrEmpty(filePath) && filePath!.IndexOf(FolderWord, StringComparison.OrdinalIgnoreCase) >= 0;

    /// <summary>
    /// The head of a file the path rule applies to: the detected head when the file has one; else line 1 as the head when it is an
    /// "&amp;Key=" line with text after it; else an EMPTY head (no metadata, ends at the first text byte: there is no head block, see
    /// <see cref="HasHeadBlock"/>). Null only for UTF-16 text (the marks are single-byte) and a file with no text.
    /// </summary>
    public static ToratEmetHead? ReadHead(ReadOnlySpan<byte> sample, int bomLength, Encoding encoding)
    {
        if (encoding.CodePage is MarkupTextEncoding.Utf16LittleEndian or MarkupTextEncoding.Utf16BigEndian) return null;
        if (bomLength >= sample.Length) return null;
        ToratEmetHead? detected = ToratEmetFormat.TryReadHead(sample, bomLength, encoding);
        if (detected is not null) return detected;
        if (sample[bomLength] == (byte)'&')
        {
            int length = Math.Min(sample.Length, bomLength + ToratEmetFormat.HeadSearchBytes);
            int end = sample.Slice(bomLength, length - bomLength).IndexOfAny((byte)'\r', (byte)'\n');
            if (end >= 0)
            {
                int next = bomLength + end;
                if (next + 1 < length && sample[next] == '\r' && sample[next + 1] == '\n') next++;
                next++;
                var pairs = ToratEmetFormat.ParseKeyValueLine(encoding.GetString(sample.Slice(bomLength, end).ToArray()));
                if (pairs.Count > 0 && next < sample.Length) return new ToratEmetHead(next, pairs);
            }
        }
        return new ToratEmetHead(bomLength, []);
    }

    /// <summary>False for the empty head of a file without metadata: block 0 is then its first text line, not a hidden head.</summary>
    public static bool HasHeadBlock(ToratEmetHead head) => head.Metadata.Count > 0;
}

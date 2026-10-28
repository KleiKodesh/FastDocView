namespace FastDocView.Markup;

public enum MimeTransferEncoding : byte { Identity = 0, QuotedPrintable = 1, Base64 = 2 }

/// <summary>
/// Streaming Content-Transfer-Encoding decoder (RFC 2045 §6.7 quoted-printable, §6.8 base64) that
/// reports, for every decoded byte, a <b>resume point</b>: where decoding can restart so that this
/// byte is the next one out. Resume point = encoded file offset × 4 + bytes to skip:
///   • identity / quoted-printable: the offset of the token that produced the byte ("=XX" or the
///     literal), skip 0 — a token start carries no decoder state (a soft line break produces nothing);
///   • base64: the offset of the 4-character group's first character, skip 0-2 — groups are the only
///     aligned restart points (RFC 2045 §6.8: 4 characters → 3 octets; non-alphabet characters ignored).
/// Quoted-printable is decoded leniently as RFC 2045 §6.7 recommends: "=" + lower-case hex is accepted,
/// an "=" not followed by hex or a line break is kept literally, whitespace after a soft-break "=" is
/// transport padding. Trailing whitespace before a hard line break is kept (harmless in HTML).
/// </summary>
public sealed class MimeTransferDecoder
{
    private enum QuotedPrintableState : byte { Normal, Equal, EqualHex, SoftCarriageReturn, SoftWhitespace }

    private readonly MimeTransferEncoding _encoding;
    private QuotedPrintableState _qpState;
    private long _tokenStart;
    private int _firstHex;
    private readonly int[] _group = new int[4];
    private int _groupCount;
    private long _groupStart;

    public MimeTransferDecoder(MimeTransferEncoding encoding) => _encoding = encoding;

    public static long ResumePoint(long encodedOffset, int skip) => encodedOffset * 4 + skip;
    public static long EncodedOffsetOf(long resumePoint) => resumePoint >> 2;
    public static int SkipOf(long resumePoint) => (int)(resumePoint & 3);

    /// <summary>Restart at an encoded offset taken from a resume point (state cleared).</summary>
    public void Reset()
    {
        _qpState = QuotedPrintableState.Normal;
        _groupCount = 0;
    }

    /// <summary>
    /// Decodes <paramref name="count"/> encoded bytes that sit at file offset <paramref name="fileOffset"/>.
    /// Output goes to <paramref name="output"/> / <paramref name="resumePoints"/> (resume array may be null)
    /// from <paramref name="outputCount"/>; both must have room for <paramref name="count"/> + 3 bytes.
    /// </summary>
    public void Feed(byte[] input, int offset, int count, long fileOffset, byte[] output, long[]? resumePoints, ref int outputCount)
    {
        switch (_encoding)
        {
            case MimeTransferEncoding.QuotedPrintable:
                for (int i = 0; i < count; i++) FeedQuotedPrintable(input[offset + i], fileOffset + i, output, resumePoints, ref outputCount);
                return;
            case MimeTransferEncoding.Base64:
                for (int i = 0; i < count; i++) FeedBase64(input[offset + i], fileOffset + i, output, resumePoints, ref outputCount);
                return;
            default:
                Buffer.BlockCopy(input, offset, output, outputCount, count);
                if (resumePoints is not null) for (int i = 0; i < count; i++) resumePoints[outputCount + i] = (fileOffset + i) * 4;
                outputCount += count;
                return;
        }
    }

    private static void Emit(byte value, long resume, byte[] output, long[]? resumePoints, ref int outputCount)
    {
        output[outputCount] = value;
        if (resumePoints is not null) resumePoints[outputCount] = resume;
        outputCount++;
    }

    private static int HexValue(byte b) => b is >= (byte)'0' and <= (byte)'9' ? b - '0' : b is >= (byte)'A' and <= (byte)'F' ? b - 'A' + 10 : b is >= (byte)'a' and <= (byte)'f' ? b - 'a' + 10 : -1;

    private void FeedQuotedPrintable(byte b, long at, byte[] output, long[]? resume, ref int count)
    {
        while (true)
        {
            switch (_qpState)
            {
                case QuotedPrintableState.Normal:
                    if (b == (byte)'=') { _qpState = QuotedPrintableState.Equal; _tokenStart = at; return; }
                    Emit(b, at * 4, output, resume, ref count);
                    return;
                case QuotedPrintableState.Equal:
                    if (HexValue(b) >= 0) { _firstHex = b; _qpState = QuotedPrintableState.EqualHex; return; }
                    if (b == (byte)'\r') { _qpState = QuotedPrintableState.SoftCarriageReturn; return; }
                    if (b == (byte)'\n') { _qpState = QuotedPrintableState.Normal; return; }
                    if (b == (byte)' ' || b == (byte)'\t') { _qpState = QuotedPrintableState.SoftWhitespace; return; }
                    Emit((byte)'=', _tokenStart * 4, output, resume, ref count);
                    _qpState = QuotedPrintableState.Normal;
                    continue;                                   // reprocess b as a literal
                case QuotedPrintableState.EqualHex:
                    if (HexValue(b) >= 0)
                    {
                        Emit((byte)(HexValue((byte)_firstHex) * 16 + HexValue(b)), _tokenStart * 4, output, resume, ref count);
                        _qpState = QuotedPrintableState.Normal;
                        return;
                    }
                    Emit((byte)'=', _tokenStart * 4, output, resume, ref count);
                    Emit((byte)_firstHex, (_tokenStart + 1) * 4, output, resume, ref count);
                    _qpState = QuotedPrintableState.Normal;
                    continue;
                case QuotedPrintableState.SoftCarriageReturn:
                    _qpState = QuotedPrintableState.Normal;
                    if (b == (byte)'\n') return;
                    continue;
                case QuotedPrintableState.SoftWhitespace:
                    if (b == (byte)' ' || b == (byte)'\t') return;
                    if (b == (byte)'\r') { _qpState = QuotedPrintableState.SoftCarriageReturn; return; }
                    if (b == (byte)'\n') { _qpState = QuotedPrintableState.Normal; return; }
                    Emit((byte)'=', _tokenStart * 4, output, resume, ref count);
                    _qpState = QuotedPrintableState.Normal;
                    continue;
            }
        }
    }

    private static int Base64Value(byte b) => b switch
    {
        >= (byte)'A' and <= (byte)'Z' => b - 'A',
        >= (byte)'a' and <= (byte)'z' => b - 'a' + 26,
        >= (byte)'0' and <= (byte)'9' => b - '0' + 52,
        (byte)'+' => 62,
        (byte)'/' => 63,
        (byte)'=' => 64,
        _ => -1,
    };

    private void FeedBase64(byte b, long at, byte[] output, long[]? resume, ref int count)
    {
        int value = Base64Value(b);
        if (value < 0) return;                                  // line breaks and anything outside the alphabet are ignored
        if (_groupCount == 0) _groupStart = at;
        _group[_groupCount++] = value;
        if (_groupCount < 4) return;
        _groupCount = 0;
        if (_group[0] == 64 || _group[1] == 64) return;          // malformed padding
        int bits = (_group[0] << 18) | (_group[1] << 12) | ((_group[2] & 63) << 6) | (_group[3] & 63);
        Emit((byte)(bits >> 16), _groupStart * 4, output, resume, ref count);
        if (_group[2] == 64) return;
        Emit((byte)(bits >> 8), _groupStart * 4 + 1, output, resume, ref count);
        if (_group[3] == 64) return;
        Emit((byte)bits, _groupStart * 4 + 2, output, resume, ref count);
    }
}

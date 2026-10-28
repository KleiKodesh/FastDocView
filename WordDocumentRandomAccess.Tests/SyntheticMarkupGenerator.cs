using System.Text;
using WordDocumentRandomAccess.Markup;

namespace WordDocumentRandomAccess.Tests;

/// <summary>
/// Synthetic .txt (read as HTML) and MHTML fixtures with Hebrew text. Hebrew is written as \u escapes.
/// The .txt prefix exercises: a heading, stray '<', an inline element open across a line break, a
/// start tag spanning two lines, an empty line, a comment spanning two lines, a heading spanning two
/// lines, script, an event handler; then <c>count</c> filler lines with an h3 every 1000.
/// </summary>
public static class SyntheticMarkupGenerator
{
    public const string Title = "\u05DB\u05D5\u05EA\u05E8\u05EA";             // heading word
    public const string Line = "\u05E9\u05D5\u05E8\u05D4";                     // "line"
    public const string Bold = "\u05DE\u05D5\u05D3\u05D2\u05E9";
    public const string After = "\u05D0\u05D7\u05E8\u05D9";
    public const string Hidden = "\u05D4\u05E2\u05E8\u05D4";
    public const string Paragraph = "\u05E4\u05E1\u05E7\u05D4";

    public static readonly string[] PrefixLines =
    [
        $"<h1>{Title} 1</h1>",                                  // block 0
        $"{Line} a < b x<3 <> 5>4 & {Line}",                    // 1: stray '<'
        $"<b>{Bold} start",                                     // 2
        $"{Bold} end</b> {Line}",                               // 3: reopens <b>
        "<div class=\"x\"",                                      // 4: tag spans lines 4-5
        $" title='y>z'>{Line} div</div>",
        "",                                                      // 5: empty
        $"<!-- {Hidden}",                                        // 6: comment spans lines
        $"{Hidden} -->{After}",
        $"<h2>{Title}",                                         // 7: heading over two lines
        $"2</h2>",                                               // 8
        $"<script>alert(1)</script>{Line} s",                   // 9
        $"<p onclick=\"x()\">{Paragraph}</p>",                  // 10
    ];

    public const int PrefixBlockCount = 11;

    public static string FillerLine(int i) => i % 1000 == 0 ? $"<h3>{Title} {i}</h3>" : $"{Line} {i} <i>{Paragraph}</i> {i}";

    public static int ExpectedBlockCount(int count) => PrefixBlockCount + count;
    public static int ExpectedHeadingCount(int count) => 2 + (count + 999) / 1000;

    public static void WriteText(string path, int count, Encoding encoding, bool withBom, string newline)
    {
        using var file = new FileStream(path, FileMode.Create, FileAccess.Write);
        if (withBom)
        {
            byte[] bom = encoding.CodePage switch { 65001 => [0xEF, 0xBB, 0xBF], 1200 => [0xFF, 0xFE], 1201 => [0xFE, 0xFF], _ => [] };
            file.Write(bom, 0, bom.Length);
        }
        using var writer = new StreamWriter(file, encoding.WebName == "utf-8" ? new UTF8Encoding(false) : NoPreamble(encoding), 1 << 16);
        writer.NewLine = newline;
        foreach (string line in PrefixLines) writer.WriteLine(line);
        for (int i = 0; i < count; i++) writer.WriteLine(FillerLine(i));
    }

    private static Encoding NoPreamble(Encoding encoding) => encoding.CodePage switch
    {
        1200 => new UnicodeEncoding(false, false),
        1201 => new UnicodeEncoding(true, false),
        _ => encoding,
    };

    // ── MHTML ────────────────────────────────────────────────────────────────────────

    public const string Boundary = "----=_NextPart_01DA0000.TEST";
    public const int MhtmlPrefixBlockCount = 4;
    public static int MhtmlExpectedBlockCount(int count) => MhtmlPrefixBlockCount + count;
    public static int MhtmlExpectedHeadingCount(int count) => 1 + (count + 499) / 500;
    public static string MhtmlParagraphText(int i) => $"{Paragraph} {i}";

    public static string WordHtml(int count)
    {
        var html = new StringBuilder();
        html.Append("<html xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\">\r\n<head>\r\n");
        html.Append("<meta http-equiv=Content-Type content=\"text/html; charset=windows-1255\">\r\n");
        html.Append("<title>").Append(Title).Append("</title>\r\n");
        html.Append("<style>\r\n<!--\r\n /* Style Definitions */\r\n p.MsoNormal, li.MsoNormal, div.MsoNormal\r\n\t{margin:0cm;\r\n\tfont-size:12.0pt;\r\n\tfont-family:\"Times New Roman\";}\r\n");
        html.Append("@page WordSection1\r\n\t{size:595.3pt 841.9pt;}\r\ndiv.WordSection1\r\n\t{page:WordSection1;}\r\n-->\r\n</style>\r\n</head>\r\n\r\n");
        html.Append("<body lang=HE dir=RTL>\r\n\r\n<div class=WordSection1 dir=RTL>\r\n\r\n");
        html.Append("<h1 dir=RTL>").Append(Title).Append(" 1</h1>\r\n\r\n");
        html.Append("<p class=MsoNormal dir=RTL>").Append(Paragraph).Append(" <b>").Append(Bold).Append("</b> a=b<o:p></o:p></p>\r\n\r\n");
        html.Append("<table class=MsoTableGrid border=1 cellspacing=0>\r\n <tr>\r\n  <td width=100>\r\n  <p class=MsoNormal>").Append(Line).Append("</p>\r\n  </td>\r\n </tr>\r\n</table>\r\n\r\n");
        html.Append("<p class=MsoNormal><!--[if gte vml 1]><v:shape id=\"x\"><v:imagedata src=\"image001.png\" o:title=\"\"/></v:shape><![endif]--><![if !vml]><img width=10 height=10 src=\"image001.png\" v:shapes=\"x\"><![endif]></p>\r\n\r\n");
        for (int i = 0; i < count; i++)
        {
            if (i % 500 == 0) html.Append("<h2 dir=RTL>").Append(Title).Append(' ').Append(i).Append("</h2>\r\n\r\n");
            else html.Append("<p class=MsoNormal dir=RTL>").Append(MhtmlParagraphText(i)).Append("</p>\r\n\r\n");
        }
        html.Append("</div>\r\n\r\n</body>\r\n\r\n</html>\r\n");
        return html.ToString();
    }

    public static void WriteMhtml(string path, int count, MimeTransferEncoding transfer)
    {
        Encoding hebrew = MarkupTextEncoding.FromCodePage(1255);
        byte[] html = hebrew.GetBytes(WordHtml(count));
        using var file = new FileStream(path, FileMode.Create, FileAccess.Write);
        void Write(string text) { byte[] b = Encoding.ASCII.GetBytes(text); file.Write(b, 0, b.Length); }
        Write("MIME-Version: 1.0\r\nX-Document-Type: Word\r\nContent-Type: multipart/related;\r\n\tboundary=\"" + Boundary + "\"\r\n\r\n");
        Write("This document is a Single File Web Page, also known as a Web Archive file.\r\n\r\n");
        Write("--" + Boundary + "\r\nContent-Location: file:///C:/doc.htm\r\n");
        Write("Content-Transfer-Encoding: " + (transfer == MimeTransferEncoding.Base64 ? "base64" : "quoted-printable") + "\r\n");
        Write("Content-Type: text/html; charset=\"windows-1255\"\r\n\r\n");
        Write(transfer == MimeTransferEncoding.Base64 ? Base64Lines(html) : QuotedPrintable(html));
        Write("\r\n--" + Boundary + "\r\nContent-Location: file:///C:/doc_files/image001.png\r\nContent-Transfer-Encoding: base64\r\nContent-Type: image/png\r\n\r\n");
        Write(Base64Lines(new byte[300]));
        Write("\r\n--" + Boundary + "--\r\n");
    }

    private static string Base64Lines(byte[] data)
    {
        string all = Convert.ToBase64String(data);
        var text = new StringBuilder();
        for (int i = 0; i < all.Length; i += 76) text.Append(all, i, Math.Min(76, all.Length - i)).Append("\r\n");
        return text.ToString();
    }

    /// <summary>RFC 2045 §6.7: '=' and 8-bit bytes as =XX, CRLF kept as hard breaks, soft breaks keep lines ≤ 76.</summary>
    private static string QuotedPrintable(byte[] data)
    {
        var text = new StringBuilder();
        int lineLength = 0;
        for (int i = 0; i < data.Length; i++)
        {
            byte b = data[i];
            if (b == '\r' && i + 1 < data.Length && data[i + 1] == '\n') { text.Append("\r\n"); lineLength = 0; i++; continue; }
            string token = b == '=' || b > 126 || (b < 32 && b != '\t') ? "=" + b.ToString("X2") : ((char)b).ToString();
            if (lineLength + token.Length > 75) { text.Append("=\r\n"); lineLength = 0; }
            text.Append(token);
            lineLength += token.Length;
        }
        return text.ToString();
    }
}

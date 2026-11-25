using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace FastDocView.Docx;

/// <summary>
/// Fonts embedded in a .docx (Word: File > Options > Save > Embed fonts in the file): fontTable.xml's w:font carries
/// w:embedRegular / w:embedBold / w:embedItalic / w:embedBoldItalic (ECMA-376 Part 1, 17.8.3.3-6), each naming an obfuscated
/// font part (.odttf, 17.8.1) by r:id with its w:fontKey GUID. Read at open from the small font table (names and part names
/// only, never the font data); a face's bytes are read when the page asks for it, through
/// <see cref="IWordWindowedDocument.ReadPicture"/> with the id <see cref="FontIdPrefix"/> + part name, de-obfuscated there.
/// The page gets the faces as @font-face rules in the document stylesheet (<see cref="StylesheetRules"/>): local() first, so an
/// installed font of that name is used as is and nothing is read, the embedded data only when the font is missing - as Word
/// draws an embedded font only when the machine lacks it. LibreOffice reads the same elements
/// (sw/source/writerfilter/dmapper/FontTable.cxx:106-118, key from the GUID at 278-294) and de-obfuscates the first 32 bytes
/// (vcl/source/gdi/embeddedfontsmanager.cxx:283-296).
/// </summary>
public sealed class DocxEmbeddedFonts
{
    /// <summary>The picture id prefix of an embedded font face: "o:" + its part name.</summary>
    public const string FontIdPrefix = "o:";

    public sealed record Face(string Family, bool IsBold, bool IsItalic, string PartName, byte[] Key);

    public static readonly DocxEmbeddedFonts None = new([]);

    public IReadOnlyList<Face> Faces { get; }
    private readonly Dictionary<string, Face> _byPart;

    private DocxEmbeddedFonts(List<Face> faces)
    {
        Faces = faces;
        _byPart = new Dictionary<string, Face>(StringComparer.Ordinal);
        foreach (Face face in faces) _byPart[face.PartName] = face;
    }

    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>The faces a font table part names (<paramref name="fontRelationships"/>: its .rels, r:id -> part name).</summary>
    public static DocxEmbeddedFonts Read(XDocument? fontTable, IReadOnlyDictionary<string, string> fontRelationships)
    {
        if (fontTable?.Root is not { } root) return None;
        var faces = new List<Face>();
        foreach (XElement font in root.Elements(W + "font"))
        {
            string? family = (string?)font.Attribute(W + "name");
            if (string.IsNullOrWhiteSpace(family)) continue;
            foreach ((string element, bool bold, bool italic) in Styles)
            {
                if (font.Element(W + element) is not { } embed) continue;
                string? id = (string?)embed.Attribute(R + "id");
                if (id is null || !fontRelationships.TryGetValue(id, out string? part)) continue;
                // without a key the part is not obfuscated (17.8.1: obfuscation applies when fontKey is given)
                byte[] key = KeyOf((string?)embed.Attribute(W + "fontKey"));
                faces.Add(new Face(family!.Trim(), bold, italic, part, key));
            }
        }
        return faces.Count == 0 ? None : new DocxEmbeddedFonts(faces);
    }

    private static readonly (string Element, bool Bold, bool Italic)[] Styles =
        [("embedRegular", false, false), ("embedBold", true, false), ("embedItalic", false, true), ("embedBoldItalic", true, true)];

    /// <summary>
    /// The 16 key bytes of a "{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}" font key: the GUID's hex digits read as bytes from its
    /// last pair to its first (ECMA-376 Part 1, 17.8.1; LibreOffice FontTable.cxx:280-293). Empty when there is no valid key.
    /// </summary>
    public static byte[] KeyOf(string? fontKey)
    {
        if (fontKey is null) return [];
        var hex = new StringBuilder(32);
        foreach (char c in fontKey)
            if (Uri.IsHexDigit(c)) hex.Append(c);
            else if (c is not ('{' or '}' or '-')) return [];
        if (hex.Length != 32) return [];
        var key = new byte[16];
        for (int i = 0; i < 16; i++)
            key[i] = byte.Parse(hex.ToString(30 - 2 * i, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        return key;
    }

    /// <summary>De-obfuscates a font part in place: its first 32 bytes XOR the key, repeated (17.8.1).</summary>
    public static void Deobfuscate(byte[] data, byte[] key)
    {
        if (key.Length != 16) return;
        for (int i = 0; i < 32 && i < data.Length; i++) data[i] ^= key[i % 16];
    }

    /// <summary>The face <paramref name="pictureId"/> names, de-obfuscated, as a font file; null for any other id.</summary>
    public WordDocumentPicture? Read(string pictureId, Func<string, byte[]?> readPart)
    {
        if (!pictureId.StartsWith(FontIdPrefix, StringComparison.Ordinal) || !_byPart.TryGetValue(pictureId.Substring(FontIdPrefix.Length), out Face? face)) return null;
        if (readPart(face.PartName) is not { } data) return null;
        Deobfuscate(data, face.Key);
        return new WordDocumentPicture(ContentTypeOf(data), data);
    }

    private static string ContentTypeOf(byte[] data)
        => data.Length >= 4 && data[0] == (byte)'O' && data[1] == (byte)'T' && data[2] == (byte)'T' && data[3] == (byte)'O' ? "font/otf"
           : data.Length >= 4 && data[0] == (byte)'t' && data[1] == (byte)'t' && data[2] == (byte)'c' && data[3] == (byte)'f' ? "font/collection"
           : "font/ttf";

    /// <summary>
    /// One @font-face rule per face: the family's installed font first (local()), then the embedded data by its picture id as
    /// <c>url("fdv-font:&lt;id&gt;")</c> - the host resolves it to its picture URL (the app registers the faces itself).
    /// </summary>
    public string StylesheetRules()
    {
        if (Faces.Count == 0) return "";
        var css = new StringBuilder();
        foreach (Face face in Faces)
        {
            string family = new string(face.Family.Where(c => c is not ('\'' or '"' or '\\' or ';' or '{' or '}' or '<' or '>') && !char.IsControl(c)).ToArray());
            // local() names a face by its full name: "Arial", "Arial Bold", "Arial Italic", "Arial Bold Italic"
            string fullName = family + (face.IsBold ? " Bold" : "") + (face.IsItalic ? " Italic" : "");
            css.Append("@font-face { font-family: '").Append(family).Append("'; font-weight: ").Append(face.IsBold ? "700" : "400")
               .Append("; font-style: ").Append(face.IsItalic ? "italic" : "normal").Append("; src: local('").Append(fullName).Append("'), url(\"fdv-font:")
               .Append(Uri.EscapeDataString(FontIdPrefix + face.PartName)).Append("\"); }\n");
        }
        return css.ToString();
    }
}

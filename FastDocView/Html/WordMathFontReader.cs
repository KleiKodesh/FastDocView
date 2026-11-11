using System.Text;

namespace FastDocView.Html;

/// <summary>
/// The stretch data of an installed math font, read at run time from its OpenType MATH table (OpenType 1.9 "MATH - The
/// Mathematical Typesetting Table": MathVariants -> vertical glyph coverage, MathGlyphConstruction with its
/// MathGlyphVariantRecords (variant glyph, advance height) and GlyphAssembly part records) plus head (unitsPerEm), cmap
/// (format 4 / 12) and hhea / hmtx (advance widths). A font is found by its family or full name (name table IDs 1 / 4 /
/// 16) among the files in the Windows font folders (system and per-user), .ttf / .otf / .ttc; the files are scanned once
/// for names, each font's table is read once and kept. Not on Windows, or a font without a file or a MATH table: null, and
/// the caller falls back to <see cref="WordMathFontVariants"/> (tables generated for the fonts Word offers).
/// </summary>
internal static class WordMathFontReader
{
    private static readonly object Lock = new();
    private static Dictionary<string, (string path, int face)>? _files;
    private static readonly Dictionary<string, Dictionary<string, (double[] heights, double[] widths, double assemblyWidth)>?> Fonts = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>The characters asked for: the delimiters and n-ary operators the converter stretches.</summary>
    private static readonly string Characters = "()[]{}|‖⟨⟩⌈⌉⌊⌋∫∬∭∮∯∰∑∏∐⋃⋂⋁⋀⨁⨂⨀⨄⨆";

    private static readonly Dictionary<string, double?> LineHeights = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>The font's line height (hhea ascender - descender + line gap, em); null when its file cannot be read.</summary>
    public static double? LineHeight(string fontName)
    {
        lock (Lock)
        {
            if (LineHeights.TryGetValue(fontName, out double? known)) return known;
            double? height = null;
            try
            {
                _files ??= ScanFontFiles();
                if (_files.TryGetValue(fontName, out var file))
                {
                    using FileStream stream = new(file.path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
                    var r = new BinaryReader(stream);
                    var tables = Tables(r, FaceOffsets(r)[Math.Max(0, file.face)]);
                    if (tables.TryGetValue("head", out var head) && tables.TryGetValue("hhea", out var hhea))
                    {
                        r.BaseStream.Position = head.offset + 18;
                        double upem = U16(r);
                        r.BaseStream.Position = hhea.offset + 4;
                        double ascent = S16(r), descent = S16(r), gap = S16(r);
                        if (upem > 0) height = (ascent - descent + gap) / upem;
                    }
                }
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or EndOfStreamException or ArgumentException or IndexOutOfRangeException) { height = null; }
            LineHeights[fontName] = height;
            return height;
        }
    }

    public static Dictionary<string, (double[] heights, double[] widths, double assemblyWidth)>? Variants(string fontName)
    {
        lock (Lock)
        {
            if (Fonts.TryGetValue(fontName, out var known)) return known;
            Dictionary<string, (double[], double[], double)>? read = null;
            try
            {
                _files ??= ScanFontFiles();
                if (_files.TryGetValue(fontName, out var file))
                    using (FileStream stream = new(file.path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                        read = ReadVariants(stream, file.face);
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or EndOfStreamException or ArgumentException or InvalidDataException or NotSupportedException or IndexOutOfRangeException) { read = null; }
            Fonts[fontName] = read;
            return read;
        }
    }

    private static IEnumerable<string> FontFolders()
    {
        string system = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        if (system.Length > 0) yield return system;
        string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (local.Length > 0) yield return Path.Combine(local, "Microsoft", "Windows", "Fonts");
    }

    /// <summary>Family / full name -> file and face, over every font file (name tables only).</summary>
    private static Dictionary<string, (string path, int face)> ScanFontFiles()
    {
        var files = new Dictionary<string, (string, int)>(StringComparer.OrdinalIgnoreCase);
        foreach (string folder in FontFolders())
        {
            if (!Directory.Exists(folder)) continue;
            foreach (string path in Directory.EnumerateFiles(folder))
            {
                string extension = Path.GetExtension(path).ToLowerInvariant();
                if (extension is not (".ttf" or ".otf" or ".ttc")) continue;
                try
                {
                    using FileStream stream = new(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
                    var reader = new BinaryReader(stream);
                    List<long> faces = FaceOffsets(reader);
                    for (int face = 0; face < faces.Count; face++)
                    {
                        var tables = Tables(reader, faces[face]);
                        if (!tables.ContainsKey("MATH") || !tables.TryGetValue("name", out var name)) continue;
                        foreach (string fontName in Names(reader, name.offset)) if (!files.ContainsKey(fontName)) files.Add(fontName, (path, face));   // net48 has no TryAdd
                    }
                }
                catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or EndOfStreamException or ArgumentException) { }
            }
        }
        return files;
    }

    private static ushort U16(BinaryReader r) { var b = r.ReadBytes(2); if (b.Length < 2) throw new EndOfStreamException(); return (ushort)(b[0] << 8 | b[1]); }
    private static short S16(BinaryReader r) => (short)U16(r);
    private static uint U32(BinaryReader r) { var b = r.ReadBytes(4); if (b.Length < 4) throw new EndOfStreamException(); return (uint)(b[0] << 24 | b[1] << 16 | b[2] << 8 | b[3]); }

    private static List<long> FaceOffsets(BinaryReader r)
    {
        r.BaseStream.Position = 0;
        uint tag = U32(r);
        if (tag != 0x74746366) return [0];                       // 'ttcf': a collection
        U32(r);
        uint count = U32(r);
        var offsets = new List<long>();
        for (uint i = 0; i < Math.Min(count, 64u); i++) offsets.Add(U32(r));
        return offsets;
    }

    private static Dictionary<string, (long offset, long length)> Tables(BinaryReader r, long faceOffset)
    {
        r.BaseStream.Position = faceOffset + 4;
        int count = U16(r);
        r.BaseStream.Position = faceOffset + 12;
        var tables = new Dictionary<string, (long, long)>(StringComparer.Ordinal);
        for (int i = 0; i < Math.Min(count, 256); i++)
        {
            string tag = Encoding.ASCII.GetString(r.ReadBytes(4));
            U32(r);
            long offset = U32(r), length = U32(r);
            tables[tag] = (offset, length);
        }
        return tables;
    }

    private static IEnumerable<string> Names(BinaryReader r, long table)
    {
        r.BaseStream.Position = table + 2;
        int count = U16(r), strings = U16(r);
        var names = new List<string>();
        for (int i = 0; i < Math.Min(count, 512); i++)
        {
            r.BaseStream.Position = table + 6 + i * 12;
            int platform = U16(r), encoding = U16(r); U16(r);
            int id = U16(r), length = U16(r), offset = U16(r);
            if (id is not (1 or 4 or 16) || platform != 3 || encoding is not (0 or 1)) continue;
            r.BaseStream.Position = table + strings + offset;
            names.Add(Encoding.BigEndianUnicode.GetString(r.ReadBytes(length)));
        }
        return names;
    }

    private static Dictionary<string, (double[], double[], double)>? ReadVariants(Stream stream, int face)
    {
        var r = new BinaryReader(stream);
        long faceOffset = FaceOffsets(r)[Math.Max(0, face)];
        var tables = Tables(r, faceOffset);
        if (!tables.TryGetValue("MATH", out var math) || !tables.TryGetValue("head", out var head) || !tables.TryGetValue("cmap", out var cmap)
            || !tables.TryGetValue("hhea", out var hhea) || !tables.TryGetValue("hmtx", out var hmtx)) return null;
        r.BaseStream.Position = head.offset + 18;
        double upem = U16(r);
        if (upem <= 0) return null;
        r.BaseStream.Position = hhea.offset + 34;
        int metrics = U16(r);
        double Advance(int glyph)
        {
            r.BaseStream.Position = hmtx.offset + Math.Min(glyph, Math.Max(0, metrics - 1)) * 4;
            return U16(r) / upem;
        }
        Dictionary<int, int> glyphs = CharacterGlyphs(r, cmap.offset);
        // MATH header: version, MathConstants, MathGlyphInfo, MathVariants (offsets from the table start)
        r.BaseStream.Position = math.offset + 8;
        long variants = math.offset + U16(r);
        r.BaseStream.Position = variants + 2;
        long vertCoverage = variants + U16(r);
        U16(r);
        int vertCount = U16(r);
        var constructionOffsets = new long[vertCount];
        r.BaseStream.Position = variants + 10;
        r.BaseStream.Position = variants + 10;
        for (int i = 0; i < vertCount; i++) constructionOffsets[i] = variants + U16(r);
        List<int> coverage = Coverage(r, vertCoverage);
        var result = new Dictionary<string, (double[], double[], double)>(StringComparer.Ordinal);
        foreach (char c in Characters)
        {
            if (!glyphs.TryGetValue(c, out int glyph)) continue;
            int index = coverage.IndexOf(glyph);
            if (index < 0 || index >= constructionOffsets.Length) continue;
            long construction = constructionOffsets[index];
            r.BaseStream.Position = construction;
            int assemblyOffset = U16(r), count = U16(r);
            var records = new List<(int glyph, int height)>();
            for (int i = 0; i < Math.Min(count, 64); i++) records.Add((U16(r), U16(r)));
            double assembly = 0;
            if (assemblyOffset != 0)
            {
                r.BaseStream.Position = construction + assemblyOffset + 4;     // italics correction MathValueRecord
                int parts = U16(r);
                var partGlyphs = new List<int>();
                for (int i = 0; i < Math.Min(parts, 32); i++) { long at = construction + assemblyOffset + 6 + i * 10; r.BaseStream.Position = at; partGlyphs.Add(U16(r)); }
                foreach (int part in partGlyphs) assembly = Math.Max(assembly, Advance(part));
            }
            result[c.ToString()] = (records.Select(x => x.height / upem).ToArray(), records.Select(x => Advance(x.glyph)).ToArray(), assembly);
        }
        return result.Count > 0 ? result : null;
    }

    private static List<int> Coverage(BinaryReader r, long at)
    {
        r.BaseStream.Position = at;
        int format = U16(r);
        var glyphs = new List<int>();
        if (format == 1) { int count = U16(r); for (int i = 0; i < count; i++) glyphs.Add(U16(r)); }
        else if (format == 2)
        {
            int ranges = U16(r);
            for (int i = 0; i < ranges; i++) { int start = U16(r), end = U16(r); U16(r); for (int g = start; g <= end && glyphs.Count < 65536; g++) glyphs.Add(g); }
        }
        return glyphs;
    }

    /// <summary>The glyphs of <see cref="Characters"/> from the Windows Unicode cmap (format 4 or 12).</summary>
    private static Dictionary<int, int> CharacterGlyphs(BinaryReader r, long cmap)
    {
        r.BaseStream.Position = cmap + 2;
        int count = U16(r);
        long best = -1;
        for (int i = 0; i < count; i++)
        {
            int platform = U16(r), encoding = U16(r);
            long offset = U32(r);
            if (platform == 3 && encoding == 10 || platform == 3 && encoding == 1 && best < 0) best = cmap + offset;
        }
        var map = new Dictionary<int, int>();
        if (best < 0) return map;
        r.BaseStream.Position = best;
        int format = U16(r);
        if (format == 4)
        {
            U16(r); U16(r);
            int segments = U16(r) / 2;
            r.BaseStream.Position = best + 14;
            var ends = new int[segments]; for (int i = 0; i < segments; i++) ends[i] = U16(r);
            U16(r);
            var starts = new int[segments]; for (int i = 0; i < segments; i++) starts[i] = U16(r);
            var deltas = new int[segments]; for (int i = 0; i < segments; i++) deltas[i] = S16(r);
            long rangeAt = r.BaseStream.Position;
            var ranges = new int[segments]; for (int i = 0; i < segments; i++) ranges[i] = U16(r);
            foreach (char c in Characters)
                for (int s = 0; s < segments; s++)
                {
                    if (c > ends[s] || c < starts[s]) continue;
                    int glyph;
                    if (ranges[s] == 0) glyph = (c + deltas[s]) & 0xFFFF;
                    else { r.BaseStream.Position = rangeAt + s * 2 + ranges[s] + (c - starts[s]) * 2; glyph = U16(r); if (glyph != 0) glyph = (glyph + deltas[s]) & 0xFFFF; }
                    if (glyph != 0) map[c] = glyph;
                    break;
                }
        }
        else if (format == 12)
        {
            U16(r); U32(r); U32(r);
            long groups = U32(r);
            for (long g = 0; g < Math.Min(groups, 100000); g++)
            {
                long start = U32(r), end = U32(r), glyph = U32(r);
                foreach (char c in Characters) if (c >= start && c <= end) map[c] = (int)(glyph + c - start);
            }
        }
        return map;
    }
}

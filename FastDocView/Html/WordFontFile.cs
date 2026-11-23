using System.Collections.Concurrent;
using System.Text;

namespace FastDocView.Html;

/// <summary>
/// Installed TrueType / OpenType font files (OpenType 1.9 "Font file": the table directory, a 'ttcf' collection header, the
/// name table), read only - nothing is written. The Windows font folders (system and per-user) and Office's cloud font cache
/// (%LocalAppData%/Microsoft/FontCache/4/CloudFonts, where Office keeps fonts it downloads on demand) are the only places looked in;
/// not on Windows they are missing and every lookup is null. A family is found by its family, typographic family or full name
/// (name IDs 1 / 16 / 4, every language) in the face of the wanted style (head.macStyle bold / italic bits): first in the files
/// Windows installs it as (<see cref="KnownFiles"/>, checked by their names and style), else in a name map of every file's name
/// and head tables, built once, lazily, by the first lookup that needs it. No lock is held across file reads: a family with a
/// known file, or one already found, never waits for the map; lookups that need the map wait only for the map itself.
/// A missing style falls back to the nearest installed one (bold italic: bold, italic, regular).
/// </summary>
internal static class WordFontFile
{
    /// <summary>Family / full name -> the file and face of each style (index: bold 1 | italic 2); built once on first need.</summary>
    private static readonly Lazy<Dictionary<string, (string Path, int Face)?[]>> NameMap = new(Scan, LazyThreadSafetyMode.ExecutionAndPublication);
    private static readonly ConcurrentDictionary<string, Lazy<(string Path, int Face)?>> Found = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>The files Windows installs common Office / Hebrew / East Asian families as (any style order): found without the name map.</summary>
    private static readonly Dictionary<string, string[]> KnownFiles = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Calibri"] = ["calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf"], ["Arial"] = ["arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf"],
        ["Times New Roman"] = ["times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf"], ["Cambria"] = ["cambria.ttc", "cambriab.ttf", "cambriai.ttf", "cambriaz.ttf"],
        ["Tahoma"] = ["tahoma.ttf", "tahomabd.ttf"], ["Segoe UI"] = ["segoeui.ttf", "segoeuib.ttf", "segoeuii.ttf", "segoeuiz.ttf"],
        ["Courier New"] = ["cour.ttf", "courbd.ttf", "couri.ttf", "courbi.ttf"], ["Verdana"] = ["verdana.ttf", "verdanab.ttf", "verdanai.ttf", "verdanaz.ttf"],
        ["Georgia"] = ["georgia.ttf", "georgiab.ttf", "georgiai.ttf", "georgiaz.ttf"],
        ["David"] = ["david.ttf", "davidbd.ttf"], ["Miriam"] = ["mriam.ttf"], ["Miriam Fixed"] = ["mriamc.ttf"], ["Narkisim"] = ["nrkis.ttf"],
        ["Gisha"] = ["gisha.ttf", "gishabd.ttf"], ["Levenim MT"] = ["lvnm.ttf", "lvnmbd.ttf"], ["FrankRuehl"] = ["frank.ttf"], ["Rod"] = ["rod.ttf"], ["Aharoni"] = ["ahronbd.ttf"],
        ["Yu Gothic"] = ["YuGothR.ttc", "YuGothB.ttc"], ["MS Gothic"] = ["msgothic.ttc"], ["Meiryo"] = ["meiryo.ttc", "meiryob.ttc"],
        ["SimSun"] = ["simsun.ttc"], ["Microsoft YaHei"] = ["msyh.ttc", "msyhbd.ttc"], ["Microsoft JhengHei"] = ["msjh.ttc", "msjhbd.ttc"], ["PMingLiU"] = ["mingliu.ttc"],
        ["Malgun Gothic"] = ["malgun.ttf", "malgunbd.ttf"],
    };

    /// <summary>The Windows font folders: system, then per-user installed fonts (read only).</summary>
    public static IEnumerable<string> Folders()
    {
        string system = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        if (system.Length > 0) yield return system;
        string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (local.Length > 0) yield return Path.Combine(local, "Microsoft", "Windows", "Fonts");
        // Office's cloud fonts (downloaded on demand, one folder per family, e.g. the Office theme's Jpan "Yu Mincho"): Word
        // draws with them, so they are measured too
        if (local.Length > 0 && Directory.Exists(Path.Combine(local, "Microsoft", "FontCache", "4", "CloudFonts")))
            foreach (string family in Directory.EnumerateDirectories(Path.Combine(local, "Microsoft", "FontCache", "4", "CloudFonts"))) yield return family;
    }

    public static FileStream Open(string path) => new(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);

    /// <summary>The file and face of an installed family's regular face; null when no font file has that name.</summary>
    public static (string Path, int Face)? Find(string family) => Find(family, false, false);

    /// <summary>The file and face of an installed family in a style (else the nearest installed style); null when no font file has that name.</summary>
    public static (string Path, int Face)? Find(string family, bool bold, bool italic)
    {
        int style = (bold ? 1 : 0) | (italic ? 2 : 0);
        return Found.GetOrAdd(style + "|" + family, _ => new Lazy<(string, int)?>(() => Locate(family, style), LazyThreadSafetyMode.ExecutionAndPublication)).Value;
    }

    private static (string Path, int Face)? Locate(string family, int style)
    {
        if (KnownFiles.TryGetValue(family, out string[]? fileNames))
        {
            var faces = new (string Path, int Face)?[4];
            foreach (string folder in Folders())
                foreach (string fileName in fileNames)
                {
                    string path = Path.Combine(folder, fileName);
                    if (File.Exists(path)) AddFaces(path, faces, family);
                }
            // found in its known files: no scan (a style Windows doesn't ship, such as FrankRuehl bold, takes the nearest face)
            if (Pick(faces, style) is { } known) return known;
        }
        return NameMap.Value.TryGetValue(family, out var styles) ? Pick(styles, style) : null;
    }

    /// <summary>The wanted style, else the nearest one (bold italic: bold, italic, regular); null when the family has no face.</summary>
    private static (string Path, int Face)? Pick((string Path, int Face)?[] faces, int style)
    {
        int[] order = style switch { 3 => [3, 1, 2, 0], 2 => [2, 0, 3, 1], 1 => [1, 0, 3, 2], _ => [0, 1, 2, 3] };
        foreach (int s in order) if (faces[s] is { } face) return face;
        return null;
    }

    /// <summary>Records each face of a file named <paramref name="family"/> under its style.</summary>
    private static void AddFaces(string path, (string Path, int Face)?[] faces, string family)
    {
        try
        {
            using FileStream stream = Open(path);
            List<long> offsets = FaceOffsets(stream);
            for (int face = 0; face < offsets.Count; face++)
            {
                var tables = Tables(stream, offsets[face]);
                if (!tables.TryGetValue("name", out var name) || !Names(Read(stream, name)).Contains(family, StringComparer.OrdinalIgnoreCase)) continue;
                faces[StyleOf(stream, tables)] ??= (path, face);
            }
        }
        catch (Exception exception) when (IsUnreadable(exception)) { }
    }

    /// <summary>head.macStyle bold (bit 0) | italic (bit 1); regular when there is no head table.</summary>
    private static int StyleOf(Stream stream, Dictionary<string, (long Offset, long Length)> tables)
        => tables.TryGetValue("head", out var head) && head.Length >= 46 ? U16(Read(stream, head), 44) & 3 : 0;

    /// <summary>Family / full name -> file and face per style over every font file of the folders (name and head tables only).</summary>
    private static Dictionary<string, (string Path, int Face)?[]> Scan()
    {
        var files = new Dictionary<string, (string Path, int Face)?[]>(StringComparer.OrdinalIgnoreCase);
        foreach (string folder in Folders())
        {
            if (!Directory.Exists(folder)) continue;
            foreach (string path in Directory.EnumerateFiles(folder))
            {
                string extension = Path.GetExtension(path).ToLowerInvariant();
                if (extension is not (".ttf" or ".otf" or ".ttc")) continue;
                try
                {
                    using FileStream stream = Open(path);
                    List<long> faces = FaceOffsets(stream);
                    for (int face = 0; face < faces.Count; face++)
                    {
                        var tables = Tables(stream, faces[face]);
                        if (!tables.TryGetValue("name", out var name)) continue;
                        int style = StyleOf(stream, tables);
                        foreach (string fontName in Names(Read(stream, name)))
                        {
                            if (!files.TryGetValue(fontName, out var styles)) files[fontName] = styles = new (string, int)?[4];
                            styles[style] ??= (path, face);
                        }
                    }
                }
                catch (Exception exception) when (IsUnreadable(exception)) { }
            }
        }
        return files;
    }

    public static bool IsUnreadable(Exception exception)
        => exception is IOException or UnauthorizedAccessException or EndOfStreamException or ArgumentException or IndexOutOfRangeException or InvalidDataException or NotSupportedException;

    /// <summary>The start of each face: one at 0, or a 'ttcf' collection's table directory offsets.</summary>
    public static List<long> FaceOffsets(Stream stream)
    {
        byte[] header = ReadAt(stream, 0, 12);
        if (U32(header, 0) != 0x74746366) return [0];
        uint count = U32(header, 8);
        byte[] offsets = ReadAt(stream, 12, (int)Math.Min(count, 64u) * 4);
        var faces = new List<long>();
        for (int i = 0; i + 3 < offsets.Length; i += 4) faces.Add(U32(offsets, i));
        return faces;
    }

    /// <summary>Tag -> (offset, length) of a face's tables.</summary>
    public static Dictionary<string, (long Offset, long Length)> Tables(Stream stream, long faceOffset)
    {
        int count = Math.Min(U16(ReadAt(stream, faceOffset + 4, 2), 0), 256);
        byte[] records = ReadAt(stream, faceOffset + 12, count * 16);
        var tables = new Dictionary<string, (long, long)>(StringComparer.Ordinal);
        for (int i = 0; i + 15 < records.Length; i += 16)
            tables[Encoding.ASCII.GetString(records, i, 4)] = (U32(records, i + 8), U32(records, i + 12));
        return tables;
    }

    public static byte[] Read(Stream stream, (long Offset, long Length) table)
    {
        if (table.Length is < 0 or > 64 * 1024 * 1024) throw new InvalidDataException("font table too large");
        return ReadAt(stream, table.Offset, (int)table.Length);
    }

    /// <summary>The family, typographic family and full names (IDs 1 / 16 / 4) in every language, Windows Unicode records.</summary>
    public static List<string> Names(byte[] name)
    {
        var names = new List<string>();
        int count = U16(name, 2), strings = U16(name, 4);
        for (int i = 0; i < Math.Min(count, 1024) && 6 + i * 12 + 12 <= name.Length; i++)
        {
            int at = 6 + i * 12;
            int platform = U16(name, at), encoding = U16(name, at + 2), id = U16(name, at + 6), length = U16(name, at + 8), offset = U16(name, at + 10);
            if (id is not (1 or 4 or 16) || platform != 3 || encoding is not (0 or 1) || strings + offset + length > name.Length) continue;
            names.Add(Encoding.BigEndianUnicode.GetString(name, strings + offset, length));
        }
        return names;
    }

    private static byte[] ReadAt(Stream stream, long offset, int length)
    {
        stream.Position = offset;
        byte[] buffer = new byte[length];
        int read = 0;
        while (read < length)
        {
            int got = stream.Read(buffer, read, length - read);
            if (got <= 0) throw new EndOfStreamException();
            read += got;
        }
        return buffer;
    }

    public static int U16(byte[] b, int at) => at + 1 < b.Length ? b[at] << 8 | b[at + 1] : throw new EndOfStreamException();
    public static int S16(byte[] b, int at) => (short)U16(b, at);
    public static uint U32(byte[] b, int at) => at + 3 < b.Length ? (uint)(b[at] << 24 | b[at + 1] << 16 | b[at + 2] << 8 | b[at + 3]) : throw new EndOfStreamException();
}

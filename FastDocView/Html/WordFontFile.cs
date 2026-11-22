using System.Text;

namespace FastDocView.Html;

/// <summary>
/// Installed TrueType / OpenType font files (OpenType 1.9 "Font file": the table directory, a 'ttcf' collection header, the
/// name table), read only - nothing is written. The Windows font folders (system and per-user) are the only places looked in;
/// not on Windows they are missing and every lookup is null. A family is found by its family, typographic family or full name
/// (name IDs 1 / 16 / 4, every language) - first in the file Windows installs it as (<see cref="KnownFiles"/>, checked by its
/// names), else by one scan of every file's name table, kept for the process. A regular face (head.macStyle without bold /
/// italic) wins over a styled one of the same family.
/// </summary>
internal static class WordFontFile
{
    private static readonly object Lock = new();
    private static Dictionary<string, (string Path, int Face, bool Regular)>? _scanned;

    /// <summary>The file Windows installs common Office / Hebrew families as: found without scanning every font file.</summary>
    private static readonly Dictionary<string, string> KnownFiles = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Calibri"] = "calibri.ttf", ["Arial"] = "arial.ttf", ["Times New Roman"] = "times.ttf", ["Cambria"] = "cambria.ttc",
        ["Tahoma"] = "tahoma.ttf", ["Segoe UI"] = "segoeui.ttf", ["Courier New"] = "cour.ttf", ["Verdana"] = "verdana.ttf",
        ["David"] = "david.ttf", ["Miriam"] = "mriam.ttf", ["Miriam Fixed"] = "mriamc.ttf", ["Narkisim"] = "nrkis.ttf",
        ["Gisha"] = "gisha.ttf", ["Levenim MT"] = "lvnm.ttf", ["FrankRuehl"] = "frank.ttf", ["Rod"] = "rod.ttf", ["Aharoni"] = "ahronbd.ttf",
    };

    /// <summary>The Windows font folders: system, then per-user installed fonts (read only).</summary>
    public static IEnumerable<string> Folders()
    {
        string system = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        if (system.Length > 0) yield return system;
        string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (local.Length > 0) yield return Path.Combine(local, "Microsoft", "Windows", "Fonts");
    }

    public static FileStream Open(string path) => new(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);

    /// <summary>The file and face of an installed family (regular face preferred); null when no font file has that name.</summary>
    public static (string Path, int Face)? Find(string family)
    {
        lock (Lock)
        {
            if (_scanned is null && KnownFiles.TryGetValue(family, out string? fileName))
                foreach (string folder in Folders())
                {
                    string path = Path.Combine(folder, fileName);
                    if (!File.Exists(path)) continue;
                    try
                    {
                        using FileStream stream = Open(path);
                        List<long> faces = FaceOffsets(stream);
                        for (int face = 0; face < faces.Count; face++)
                        {
                            var tables = Tables(stream, faces[face]);
                            if (tables.TryGetValue("name", out var name) && Names(Read(stream, name)).Contains(family, StringComparer.OrdinalIgnoreCase)) return (path, face);
                        }
                    }
                    catch (Exception exception) when (IsUnreadable(exception)) { }
                }
            _scanned ??= Scan();
            return _scanned.TryGetValue(family, out var found) ? (found.Path, found.Face) : null;
        }
    }

    /// <summary>Family / full name -> file and face over every font file of the folders (name and head tables only).</summary>
    private static Dictionary<string, (string, int, bool)> Scan()
    {
        var files = new Dictionary<string, (string Path, int Face, bool Regular)>(StringComparer.OrdinalIgnoreCase);
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
                        bool regular = !tables.TryGetValue("head", out var head) || head.Length < 46 || (U16(Read(stream, head), 44) & 3) == 0;
                        foreach (string fontName in Names(Read(stream, name)))
                            if (!files.TryGetValue(fontName, out var known) || regular && !known.Regular) files[fontName] = (path, face, regular);
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

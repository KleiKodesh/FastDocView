namespace FastDocView;

/// <summary>
/// Bookmarks of a document in body order: each name and the block it marks (index-time data, stored in the index).
/// Lookup returns the first bookmark of a name, as Word resolves duplicates. Shared by the .odt, .rtf and .doc indexes
/// (.docx keeps its own copy in DocxRandomAccessIndex).
/// </summary>
public sealed class WordDocumentBookmarkTable
{
    public static readonly WordDocumentBookmarkTable Empty = new([], []);

    public string[] Names { get; }
    public int[] Blocks { get; }
    private Dictionary<string, int>? _lookup;

    public WordDocumentBookmarkTable(string[] names, int[] blocks)
    {
        if (names.Length != blocks.Length) throw new ArgumentException("Bookmark names and blocks differ in length.");
        Names = names;
        Blocks = blocks;
    }

    /// <summary>The block of the first bookmark called <paramref name="name"/>, or null.</summary>
    public int? BlockOf(string name)
    {
        if (_lookup is null)
        {
            var lookup = new Dictionary<string, int>(Names.Length, StringComparer.Ordinal);
            for (int i = 0; i < Names.Length; i++)
                if (!lookup.ContainsKey(Names[i])) lookup[Names[i]] = Blocks[i];
            _lookup = lookup;
        }
        return _lookup.TryGetValue(name, out int block) ? block : null;
    }

    public void Write(BinaryWriter writer)
    {
        writer.Write(Names.Length);
        foreach (string name in Names) writer.Write(name);
        IndexBinaryArrays.Write(writer, Blocks);
    }

    /// <summary>Reads a table written by <see cref="Write"/>; every block must be below <paramref name="blockCount"/>.</summary>
    public static WordDocumentBookmarkTable Read(BinaryReader reader, int blockCount)
    {
        var names = new string[reader.ReadInt32()];
        for (int i = 0; i < names.Length; i++) names[i] = reader.ReadString();
        int[] blocks = IndexBinaryArrays.Read<int>(reader);
        if (blocks.Length != names.Length || blocks.Any(block => block < 0 || block >= Math.Max(1, blockCount)))
            throw new InvalidDataException("Index: bookmark table does not match the blocks.");
        return new WordDocumentBookmarkTable(names, blocks);
    }
}

/// <summary>
/// Bookmarks found so far by a scanner that is still running (on-demand index): folds the scanner's growing lists into a
/// lookup as they grow; a bookmark counts only once its block is known.
/// </summary>
internal sealed class WordDocumentBookmarkScanLookup
{
    private readonly List<string> _names;
    private readonly List<int> _blocks;
    private readonly Dictionary<string, int> _known = new(StringComparer.Ordinal);
    private int _seen;

    public WordDocumentBookmarkScanLookup(List<string> names, List<int> blocks)
    {
        _names = names;
        _blocks = blocks;
    }

    public int? Find(string name, int knownBlockCount)
    {
        for (; _seen < _names.Count; _seen++)
            if (!_known.ContainsKey(_names[_seen])) _known[_names[_seen]] = _blocks[_seen];
        return _known.TryGetValue(name, out int block) && block < knownBlockCount ? block : null;
    }
}

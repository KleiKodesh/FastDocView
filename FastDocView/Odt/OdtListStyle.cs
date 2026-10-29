using System.Globalization;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using FastDocView.DocxHtml;

namespace FastDocView.Odt;

/// <summary>
/// One text:list-style: per level (1-10) a number or bullet definition. Attribute set and defaults
/// as LibreOffice reads them (xmloff/source/style/xmlnumi.cxx SvxXMLListLevelStyleContext_Impl):
/// text:level, style:num-format (default "1"), style:num-prefix, style:num-suffix,
/// text:start-value (default 1), text:display-levels (default 1), text:bullet-char.
/// </summary>
public sealed class OdtListStyle
{
    public const int MaxLevels = 10;

    public sealed class Level
    {
        public bool IsBullet;
        public string BulletCharacter = "\u2022";
        public string NumberFormat = "1";
        public string Prefix = "";
        public string Suffix = "";
        public int StartValue = 1;
        public int DisplayLevels = 1;
        /// <summary>
        /// The level's label-alignment geometry (ODF 1.2 19.880 style:list-level-properties with
        /// text:list-level-position-and-space-mode="label-alignment", 19.881 style:list-level-label-alignment): fo:margin-left,
        /// fo:text-indent, text:list-tab-stop-position (twips) and text:label-followed-by, as
        /// "margin,indent,tab,follow" (follow 0 listtab, 1 space, 2 nothing; tab -1 when absent); "" in the older
        /// label-width-and-position mode, which is not modelled.
        /// </summary>
        public string Layout = "";
    }

    private readonly Level[] _levels = new Level[MaxLevels];

    private OdtListStyle()
    {
        for (int i = 0; i < MaxLevels; i++) _levels[i] = new Level();
    }

    public OdtListStyle(XElement listStyle) : this()
    {
        XNamespace text = OdtStyleCascade.Text, style = OdtStyleCascade.Style, fo = OdtStyleCascade.Fo;
        foreach (XElement element in listStyle.Elements())
        {
            XElement? properties = element.Element(style + "list-level-properties");
            XElement? alignment = properties?.Element(style + "list-level-label-alignment");
            string layout = LayoutOf((string?)properties?.Attribute(text + "list-level-position-and-space-mode"), alignment is not null,
                                     (string?)alignment?.Attribute(fo + "margin-left"), (string?)alignment?.Attribute(fo + "text-indent"),
                                     (string?)alignment?.Attribute(text + "list-tab-stop-position"), (string?)alignment?.Attribute(text + "label-followed-by"));
            AddLevel(element.Name.LocalName, element.Name.Namespace == text, (local, isStyle) => (string?)element.Attribute((isStyle ? style : text) + local), layout);
        }
    }

    /// <summary>
    /// A text:list-style read straight from an XmlReader positioned on its start element (styles.xml pass: no DOM - a
    /// converted Word file has thousands of them); the reader is left after its end element.
    /// </summary>
    internal static OdtListStyle Read(XmlReader reader)
    {
        string textNs = OdtStyleCascade.Text.NamespaceName, styleNs = OdtStyleCascade.Style.NamespaceName, foNs = OdtStyleCascade.Fo.NamespaceName;
        var listStyle = new OdtListStyle();
        if (reader.IsEmptyElement) { reader.Read(); return listStyle; }
        int depth = reader.Depth;
        reader.Read();
        while (!reader.EOF && reader.Depth > depth)
        {
            if (reader.NodeType != XmlNodeType.Element || reader.Depth != depth + 1) { reader.Read(); continue; }
            // a level element: its attributes, then its properties' label-alignment geometry
            string local = reader.LocalName;
            bool isText = reader.NamespaceURI == textNs;
            var attributes = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (string name in new[] { "level", "num-prefix", "num-suffix", "num-format", "start-value", "display-levels", "bullet-char" })
            {
                bool isStyle = name is "num-prefix" or "num-suffix" or "num-format";
                if (reader.GetAttribute(name, isStyle ? styleNs : textNs) is { } value) attributes[(isStyle ? "s:" : "t:") + name] = value;
            }
            string? mode = null, margin = null, indent = null, tab = null, follow = null;
            bool hasAlignment = false;
            int levelDepth = reader.Depth;
            if (reader.IsEmptyElement) reader.Read();
            else
            {
                reader.Read();
                while (!reader.EOF && reader.Depth > levelDepth)
                {
                    if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == styleNs)
                    {
                        if (reader.LocalName == "list-level-properties" && reader.Depth == levelDepth + 1)
                            mode = reader.GetAttribute("list-level-position-and-space-mode", textNs);
                        else if (reader.LocalName == "list-level-label-alignment" && reader.Depth == levelDepth + 2)
                        {
                            hasAlignment = true;
                            margin = reader.GetAttribute("margin-left", foNs);
                            indent = reader.GetAttribute("text-indent", foNs);
                            tab = reader.GetAttribute("list-tab-stop-position", textNs);
                            follow = reader.GetAttribute("label-followed-by", textNs);
                        }
                    }
                    reader.Read();
                }
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == levelDepth) reader.Read();
            }
            listStyle.AddLevel(local, isText, (name, isStyle) => attributes.TryGetValue((isStyle ? "s:" : "t:") + name, out string? value) ? value : null,
                               LayoutOf(mode, hasAlignment, margin, indent, tab, follow));
        }
        if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) reader.Read();
        return listStyle;
    }

    /// <summary>
    /// A text:list-style parsed straight from its UTF-8 bytes (<paramref name="element"/> = the whole element, start tag
    /// to end tag), with the prefixes the part's root declares for the text, style and fo namespaces. Used by
    /// <see cref="OdtStylesListStyleFilterStream"/>: a converted Word file has thousands of list styles, and even a
    /// streaming XmlReader over them costs a few hundred ms; the element shape is fixed (ODF 1.3 part 3 16.32 text:list-style:
    /// level elements, each with style:list-level-properties / style:list-level-label-alignment), so a tag-and-attribute
    /// walk suffices. Attribute values are entity-decoded.
    /// </summary>
    internal static OdtListStyle Parse(ReadOnlySpan<byte> element, string textPrefix, string stylePrefix, string foPrefix, ParsedLevels cache, out string? name)
    {
        ParseNames names = ParseNames.For(textPrefix, stylePrefix, foPrefix);
        var listStyle = new OdtListStyle();
        var tag = new string?[ParseNames.SlotCount];
        // the list-style start tag: its style:name
        name = null;
        int position = element.IndexOf((byte)'<');
        if (position < 0) return listStyle;
        int nameEnd = position + 1;
        while (nameEnd < element.Length && element[nameEnd] is not ((byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n' or (byte)'/' or (byte)'>')) nameEnd++;
        position = ReadAttributes(element, nameEnd, names, tag);
        name = tag[ParseNames.Name];
        // each level element: from its start tag to the next level's (or the list style's end tag). A converted Word file
        // repeats the same few level definitions thousands of times: a level parsed once is reused for equal bytes.
        int levelStart = NextLevelStart(element, position, names);
        while (levelStart >= 0)
        {
            int next = NextLevelStart(element, levelStart + 1, names);
            ReadOnlySpan<byte> segment = element.Slice(levelStart, (next < 0 ? element.Length : next) - levelStart);
            if (cache.Find(segment) is not { } built)
            {
                built = ParseLevel(segment, textPrefix, names);
                cache.Add(segment, built);
            }
            if (built.index >= 0) listStyle._levels[built.index] = built.level;
            levelStart = next;
        }
        return listStyle;
    }

    private static int NextLevelStart(ReadOnlySpan<byte> element, int from, ParseNames names)
    {
        if (from >= element.Length) return -1;
        int found = element.Slice(from).IndexOf(names.LevelStartMarker);
        return found < 0 ? -1 : from + found;
    }

    /// <summary>One level element (its start tag, style:list-level-properties, style:list-level-label-alignment); index -1 without text:level.</summary>
    private static (int index, Level level) ParseLevel(ReadOnlySpan<byte> segment, string textPrefix, ParseNames names)
    {
        var level = new string?[ParseNames.SlotCount];
        var tag = new string?[ParseNames.SlotCount];
        string? levelLocalName = null;
        int position = 0;
        while (true)
        {
            int open = segment.Slice(position).IndexOf((byte)'<');
            if (open < 0) break;
            position += open + 1;
            if (position >= segment.Length) break;
            if (segment[position] is (byte)'/' or (byte)'!' or (byte)'?')
            {
                int close = segment.Slice(position).IndexOf((byte)'>');
                if (close < 0) break;
                position += close + 1;
                continue;
            }
            int nameEnd = position;
            while (nameEnd < segment.Length && segment[nameEnd] is not ((byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n' or (byte)'/' or (byte)'>')) nameEnd++;
            ReadOnlySpan<byte> qualifiedName = segment.Slice(position, nameEnd - position);
            Array.Clear(tag, 0, tag.Length);
            position = ReadAttributes(segment, nameEnd, names, tag);
            if (levelLocalName is null && qualifiedName.StartsWith(names.LevelPrefix))
            {
                levelLocalName = Encoding.UTF8.GetString(qualifiedName.Slice(textPrefix.Length + 1).ToArray());
                for (int i = 0; i < ParseNames.Mode; i++) level[i] = tag[i];
            }
            else if (levelLocalName is not null && qualifiedName.SequenceEqual(names.Properties)) level[ParseNames.Mode] = tag[ParseNames.Mode];
            else if (levelLocalName is not null && qualifiedName.SequenceEqual(names.Alignment))
            {
                level[ParseNames.HasAlignment] = "";
                for (int i = ParseNames.MarginLeft; i <= ParseNames.FollowedBy; i++) level[i] = tag[i];
            }
        }
        if (levelLocalName is null) return (-1, new Level());
        string layout = LayoutOf(level[ParseNames.Mode], level[ParseNames.HasAlignment] is not null, level[ParseNames.MarginLeft],
                                 level[ParseNames.TextIndent], level[ParseNames.TabStop], level[ParseNames.FollowedBy]);
        return BuildLevel(levelLocalName, true, (local, isStyle) => ParseNames.SlotOf(local, isStyle) is { } slot ? level[slot] : null, layout) ?? (-1, new Level());
    }

    /// <summary>
    /// Levels parsed so far, by their bytes (one cache per styles.xml pass). Levels are never changed after they are built,
    /// so equal definitions share one <see cref="Level"/>.
    /// </summary>
    internal sealed class ParsedLevels
    {
        private readonly Dictionary<int, List<(byte[] bytes, int index, Level level)>> _byHash = new();

        public (int index, Level level)? Find(ReadOnlySpan<byte> segment)
        {
            if (!_byHash.TryGetValue(Hash(segment), out List<(byte[] bytes, int index, Level level)>? candidates)) return null;
            foreach ((byte[] bytes, int index, Level level) in candidates)
                if (segment.SequenceEqual(bytes)) return (index, level);
            return null;
        }

        public void Add(ReadOnlySpan<byte> segment, (int index, Level level) built)
        {
            int hash = Hash(segment);
            if (!_byHash.TryGetValue(hash, out List<(byte[] bytes, int index, Level level)>? candidates)) _byHash[hash] = candidates = new List<(byte[], int, Level)>(1);
            candidates.Add((segment.ToArray(), built.index, built.level));
        }

        private static int Hash(ReadOnlySpan<byte> bytes)
        {
            // FNV-1a over 8-byte words: segments are a few hundred bytes
            ulong hash = 14695981039346656037;
            int i = 0;
            for (; i + 8 <= bytes.Length; i += 8) hash = (hash ^ System.Runtime.InteropServices.MemoryMarshal.Read<ulong>(bytes.Slice(i))) * 1099511628211;
            for (; i < bytes.Length; i++) hash = (hash ^ bytes[i]) * 1099511628211;
            return (int)(hash ^ (hash >> 32)) ^ bytes.Length;
        }
    }

    /// <summary>The qualified names <see cref="Parse"/> looks for, as UTF-8 bytes, for one set of prefixes.</summary>
    private sealed class ParseNames
    {
        // attribute slots; text:bullet-char shares slot 0 with style:name (a level element has no style:name)
        public const int Name = 0, Mode = 7, MarginLeft = 8, TextIndent = 9, TabStop = 10, FollowedBy = 11, HasAlignment = 12, SlotCount = 13;
        private static readonly (bool isStyle, bool isFo, string local)[] Slots =
        [
            (true, false, "name"), (false, false, "level"), (true, false, "num-prefix"), (true, false, "num-suffix"), (true, false, "num-format"),
            (false, false, "start-value"), (false, false, "display-levels"),
            (false, false, "list-level-position-and-space-mode"), (false, true, "margin-left"), (false, true, "text-indent"),
            (false, false, "list-tab-stop-position"), (false, false, "label-followed-by"),
        ];
        private static ParseNames? _last;

        private readonly string _key;
        public readonly byte[] LevelPrefix, LevelStartMarker, Properties, Alignment, BulletCharacter;
        public readonly byte[][] SlotNames;

        private ParseNames(string key, string text, string style, string fo)
        {
            _key = key;
            LevelPrefix = Encoding.UTF8.GetBytes(text + ":list-level-style-");
            LevelStartMarker = Encoding.UTF8.GetBytes("<" + text + ":list-level-style-");
            Properties = Encoding.UTF8.GetBytes(style + ":list-level-properties");
            Alignment = Encoding.UTF8.GetBytes(style + ":list-level-label-alignment");
            BulletCharacter = Encoding.UTF8.GetBytes(text + ":bullet-char");
            SlotNames = Slots.Select(slot => Encoding.UTF8.GetBytes((slot.isStyle ? style : slot.isFo ? fo : text) + ":" + slot.local)).ToArray();
        }

        public static ParseNames For(string text, string style, string fo)
        {
            string key = text + "|" + style + "|" + fo;
            ParseNames? last = _last;
            if (last is not null && last._key == key) return last;
            return _last = new ParseNames(key, text, style, fo);
        }

        /// <summary>The slot of a level attribute BuildLevel asks for.</summary>
        public static int? SlotOf(string local, bool isStyle) => (local, isStyle) switch
        {
            ("level", false) => 1, ("num-prefix", true) => 2, ("num-suffix", true) => 3, ("num-format", true) => 4,
            ("start-value", false) => 5, ("display-levels", false) => 6, ("bullet-char", false) => Name,
            _ => null,
        };
    }

    /// <summary>Reads the wanted attributes of a tag (from <paramref name="start"/> to its '&gt;') into their slots; returns the position after it.</summary>
    private static int ReadAttributes(ReadOnlySpan<byte> element, int start, ParseNames names, string?[] slots)
    {
        int position = start;
        while (position < element.Length)
        {
            byte b = element[position];
            if (b == (byte)'>') return position + 1;
            if (b is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n' or (byte)'/') { position++; continue; }
            int equals = element.Slice(position).IndexOf((byte)'=');
            if (equals < 0) return element.Length;
            int nameLength = equals;
            while (nameLength > 0 && element[position + nameLength - 1] == (byte)' ') nameLength--;      // net48: no span TrimEnd for bytes
            ReadOnlySpan<byte> attributeName = element.Slice(position, nameLength);
            position += equals + 1;
            while (position < element.Length && element[position] is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n') position++;
            if (position >= element.Length) return element.Length;
            byte quote = element[position];
            int close = element.Slice(position + 1).IndexOf(quote);
            if (close < 0) return element.Length;
            int slot = -1;
            if (attributeName.SequenceEqual(names.BulletCharacter)) slot = ParseNames.Name;
            else
                for (int i = 0; i < names.SlotNames.Length; i++)
                    if (attributeName.SequenceEqual(names.SlotNames[i])) { slot = i; break; }
            if (slot >= 0)
            {
                string value = Encoding.UTF8.GetString(element.Slice(position + 1, close).ToArray());
                slots[slot] = value.IndexOf('&') >= 0 ? System.Net.WebUtility.HtmlDecode(value) : value;
            }
            position += close + 2;
        }
        return position;
    }

    /// <param name="attribute">(local name, is style: namespace else text:) -> value.</param>
    private void AddLevel(string localName, bool isTextNamespace, Func<string, bool, string?> attribute, string layout)
    {
        if (BuildLevel(localName, isTextNamespace, attribute, layout) is { } built) _levels[built.index] = built.level;
    }

    private static (int index, Level level)? BuildLevel(string localName, bool isTextNamespace, Func<string, bool, string?> attribute, string layout)
    {
        if (!int.TryParse(attribute("level", false), NumberStyles.Integer, CultureInfo.InvariantCulture, out int levelNumber)) return null;
        int index = Math.Max(1, Math.Min(MaxLevels, levelNumber)) - 1;
        var level = new Level { Prefix = attribute("num-prefix", true) ?? "", Suffix = attribute("num-suffix", true) ?? "" };
        if (isTextNamespace && localName == "list-level-style-number")
        {
            level.NumberFormat = attribute("num-format", true) ?? "1";
            if (int.TryParse(attribute("start-value", false), NumberStyles.Integer, CultureInfo.InvariantCulture, out int start)) level.StartValue = start < 0 ? 1 : start;
            if (int.TryParse(attribute("display-levels", false), NumberStyles.Integer, CultureInfo.InvariantCulture, out int display)) level.DisplayLevels = Math.Max(1, display);
        }
        else if (isTextNamespace && localName == "list-level-style-bullet")
        {
            level.IsBullet = true;
            string bullet = attribute("bullet-char", false) ?? "";
            // Private-use symbol-font bullets (Wingdings/Symbol) have no glyph without their font.
            level.BulletCharacter = bullet.Length == 0 || bullet[0] is >= '\uE000' and <= '\uF8FF' ? "\u2022" : bullet;
        }
        else { level.IsBullet = true; level.BulletCharacter = "\u2022"; }   // list-level-style-image: a plain bullet
        level.Layout = layout;
        return (index, level);
    }

    private static string LayoutOf(string? mode, bool hasAlignment, string? marginLeft, string? textIndent, string? tabStop, string? followedBy)
    {
        if (mode != "label-alignment" || !hasAlignment) return "";
        int margin = marginLeft is null ? 0 : OdtStyleCascade.Twips(marginLeft) ?? 0;
        int indent = textIndent is null ? 0 : OdtStyleCascade.Twips(textIndent) ?? 0;
        int tab = tabStop is null ? -1 : OdtStyleCascade.Twips(tabStop) ?? -1;
        int follow = followedBy switch { "space" => 1, "nothing" => 2, _ => 0 };
        return string.Join(",", margin, indent, tab, follow);
    }

    public Level LevelAt(int level) => _levels[Math.Max(0, Math.Min(MaxLevels - 1, level))];

    /// <summary>
    /// The label of an item at <paramref name="level"/> given the current counters of its list:
    /// the last display-levels numbers joined with ".", wrapped in this level's prefix and suffix.
    /// </summary>
    public string FormatLabel(int level, int[] counters)
    {
        Level current = LevelAt(level);
        if (current.IsBullet) return current.Prefix + current.BulletCharacter + current.Suffix;
        var label = new StringBuilder(current.Prefix);
        int first = Math.Max(0, level - current.DisplayLevels + 1);
        for (int i = first; i <= level; i++)
        {
            Level part = LevelAt(i);
            if (part.IsBullet) continue;
            string number = FormatNumber(counters[i] != OdtListNumbering.Unset ? counters[i] : part.StartValue, part.NumberFormat);
            if (number.Length == 0) continue;
            if (label.Length > current.Prefix.Length) label.Append('.');
            label.Append(number);
        }
        return label.Append(current.Suffix).ToString();
    }

    /// <summary>
    /// style:num-format → text. "1 a A i I" and "" per xmloff/source/core/xmluconv.cxx
    /// convertNumFormat; longer names come from the numbering provider
    /// (i18npool defaultnumberingprovider.cxx): "א, י, ק, ..." is NUMBER_HEBREW (gematria).
    /// "א, ב, ג, ..." (CHARS_HEBREW, letters) is unverified. Unknown → decimal, as LibreOffice falls back to ARABIC.
    /// </summary>
    public static string FormatNumber(int value, string format) => format switch
    {
        "" => "",
        "1" => value.ToString(CultureInfo.InvariantCulture),
        "a" => DocxNumberingDefinitions.FormatNumber(value, "lowerLetter"),
        "A" => DocxNumberingDefinitions.FormatNumber(value, "upperLetter"),
        "i" => DocxNumberingDefinitions.FormatNumber(value, "lowerRoman"),
        "I" => DocxNumberingDefinitions.FormatNumber(value, "upperRoman"),
        _ when format.StartsWith("א, י", StringComparison.Ordinal) => DocxNumberingDefinitions.FormatNumber(value, "hebrew1"),
        _ when format.StartsWith("א, ב", StringComparison.Ordinal) => DocxNumberingDefinitions.FormatNumber(value, "hebrew2"),
        _ => value.ToString(CultureInfo.InvariantCulture),
    };
}

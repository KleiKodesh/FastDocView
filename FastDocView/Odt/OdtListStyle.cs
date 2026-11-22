using System.Threading;
using System.Globalization;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using FastDocView.DocxHtml;
using FastDocView.Html;

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
    /// <summary>The name text:outline-style (heading numbering) is kept under among the list styles; no style name holds U+0001.</summary>
    public const string OutlineStyleName = "outline";

    /// <summary>A label without its level's prefix and suffix (LibreOffice MakeNumString(..., bInclStrings false)).</summary>
    public string WithoutPrefixAndSuffix(int level, string label)
    {
        Level current = LevelAt(level);
        if (label.Length >= current.Prefix.Length + current.Suffix.Length && label.StartsWith(current.Prefix, StringComparison.Ordinal) && label.EndsWith(current.Suffix, StringComparison.Ordinal))
            return label.Substring(current.Prefix.Length, label.Length - current.Prefix.Length - current.Suffix.Length);
        return label;
    }

    public sealed class Level
    {
        public bool IsBullet;
        public string BulletCharacter = "\u2022";
        public string NumberFormat = "1";
        public string Prefix = "";
        public string Suffix = "";
        public int StartValue = 1;
        public int DisplayLevels = 1;
        /// <summary>A bullet font given by style:font-name until <see cref="ResolveFontNames"/> looks it up (not saved: list styles are resolved before the index is written).</summary>
        internal UnresolvedBulletFont? Unresolved;
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
            XElement? textProperties = element.Element(style + "text-properties");
            var extras = new LevelExtras((string?)element.Attribute(text + "style-name"), (string?)element.Attribute(Xlink + "href"),
                                         BulletFont((string?)textProperties?.Attribute(fo + "font-family")), (string?)textProperties?.Attribute(style + "font-name"),
                                         (string?)properties?.Attribute(fo + "width"), (string?)properties?.Attribute(fo + "height"));
            AddLevel(element.Name.LocalName, element.Name.Namespace == text, (local, isStyle) => (string?)element.Attribute((isStyle ? style : text) + local), layout, extras);
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
            foreach (string name in new[] { "level", "num-prefix", "num-suffix", "num-format", "num-letter-sync", "start-value", "display-levels", "bullet-char" })
            {
                bool isStyle = name is "num-prefix" or "num-suffix" or "num-format" or "num-letter-sync";
                if (reader.GetAttribute(name, isStyle ? styleNs : textNs) is { } value) attributes[(isStyle ? "s:" : "t:") + name] = value;
            }
            string? mode = null, margin = null, indent = null, tab = null, follow = null;
            string? labelStyle = reader.GetAttribute("style-name", textNs), href = reader.GetAttribute("href", Xlink.NamespaceName);
            string? fontFamily = null, fontName = null, width = null, height = null;
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
                        {
                            mode = reader.GetAttribute("list-level-position-and-space-mode", textNs);
                            width = reader.GetAttribute("width", foNs);
                            height = reader.GetAttribute("height", foNs);
                        }
                        else if (reader.LocalName == "text-properties" && reader.Depth == levelDepth + 1)
                        {
                            fontFamily = BulletFont(reader.GetAttribute("font-family", foNs));
                            fontName = reader.GetAttribute("font-name", styleNs);
                        }
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
                               LayoutOf(mode, hasAlignment, margin, indent, tab, follow), new LevelExtras(labelStyle, href, fontFamily, fontName, width, height));
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
    /// walk suffices. Attribute values are decoded as the XML parse reads them (WordDocumentXml.DecodeAttributeValue).
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
        // a converted Word file repeats whole list styles under new names (the sample: 3,058 list styles, 467 different bodies):
        // equal bytes after the start tag (which only names it) give an equal list style, so the parsed one is shared
        ReadOnlySpan<byte> body = element.Slice(Math.Min(position, element.Length));
        if (cache.FindListStyle(body) is { } parsed) return parsed;
        // the levels are parsed when the list style is first used: a converted Word file declares thousands, a window uses a few
        listStyle._unparsedBody = cache.AddListStyle(body, listStyle);
        listStyle._unparsedNames = names;
        listStyle._unparsedLevels = cache;
        listStyle._unparsedTextPrefix = textPrefix;
        return listStyle;
    }

    // set by Parse until the levels are read (EnsureLevelsParsed): the bytes after the start tag, and what parsing them needs
    private byte[]? _unparsedBody;
    private ParseNames? _unparsedNames;
    private ParsedLevels? _unparsedLevels;
    private string? _unparsedTextPrefix;

    /// <summary>Reads the levels of a list style Parse left unparsed. Under the shared lock: ParsedLevels is shared between list styles.</summary>
    private void EnsureLevelsParsed()
    {
        if (Volatile.Read(ref _unparsedBody) is null) return;
        lock (FontNameResolution)
        {
            if (_unparsedBody is not { } body) return;
            ParseNames names = _unparsedNames!;
            ParsedLevels cache = _unparsedLevels!;
            // each level element: from its start tag to the next level's (or the list style's end tag). A converted Word file
            // repeats the same few level definitions thousands of times: a level parsed once is reused for equal bytes.
            int levelStart = NextLevelStart(body, 0, names);
            while (levelStart >= 0)
            {
                int next = NextLevelStart(body, levelStart + 1, names);
                ReadOnlySpan<byte> segment = body.AsSpan(levelStart, (next < 0 ? body.Length : next) - levelStart);
                if (cache.Find(segment) is not { } built)
                {
                    built = ParseLevel(segment, _unparsedTextPrefix!, names);
                    cache.Add(segment, built);
                }
                if (built.index >= 0) _levels[built.index] = built.level;
                levelStart = next;
            }
            _unparsedNames = null;
            _unparsedLevels = null;
            _unparsedTextPrefix = null;
            Volatile.Write(ref _unparsedBody, null);
        }
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
                level[ParseNames.StyleName] = tag[ParseNames.StyleName];
                level[ParseNames.Href] = tag[ParseNames.Href];
                level[ParseNames.LetterSync] = tag[ParseNames.LetterSync];      // style:num-letter-sync (AA, BB... rather than AA, AB...)
            }
            else if (levelLocalName is not null && qualifiedName.SequenceEqual(names.Properties))
            {
                level[ParseNames.Mode] = tag[ParseNames.Mode];
                level[ParseNames.Width] = tag[ParseNames.Width];
                level[ParseNames.Height] = tag[ParseNames.Height];
            }
            else if (levelLocalName is not null && qualifiedName.SequenceEqual(names.TextProperties))
            {
                level[ParseNames.FontFamily] = BulletFont(tag[ParseNames.FontFamily]);
                level[ParseNames.FontName] = tag[ParseNames.FontName];
            }
            else if (levelLocalName is not null && qualifiedName.SequenceEqual(names.Alignment))
            {
                level[ParseNames.HasAlignment] = "";
                for (int i = ParseNames.MarginLeft; i <= ParseNames.FollowedBy; i++) level[i] = tag[i];
            }
        }
        if (levelLocalName is null) return (-1, new Level());
        string layout = LayoutOf(level[ParseNames.Mode], level[ParseNames.HasAlignment] is not null, level[ParseNames.MarginLeft],
                                 level[ParseNames.TextIndent], level[ParseNames.TabStop], level[ParseNames.FollowedBy]);
        var extras = new LevelExtras(level[ParseNames.StyleName], level[ParseNames.Href], level[ParseNames.FontFamily], level[ParseNames.FontName], level[ParseNames.Width], level[ParseNames.Height]);
        return BuildLevel(levelLocalName, true, (local, isStyle) => ParseNames.SlotOf(local, isStyle) is { } slot ? level[slot] : null, layout, extras) ?? (-1, new Level());
    }

    /// <summary>
    /// Levels and whole list styles parsed so far, by their bytes (one cache per styles.xml pass). Levels are only changed by
    /// ResolveFontNames, which gives equal bytes equal results, so equal definitions share one <see cref="Level"/> / list style.
    /// </summary>
    internal sealed class ParsedLevels
    {
        private readonly Dictionary<int, List<(byte[] bytes, int index, Level level)>> _byHash = new();
        private readonly Dictionary<int, List<(byte[] body, OdtListStyle style)>> _listStylesByHash = new();

        /// <summary>A list style parsed before from the same bytes after its start tag (only its name differs); shared, like levels.</summary>
        public OdtListStyle? FindListStyle(ReadOnlySpan<byte> body)
        {
            if (!_listStylesByHash.TryGetValue(Hash(body), out List<(byte[] body, OdtListStyle style)>? candidates)) return null;
            foreach ((byte[] bytes, OdtListStyle style) in candidates)
                if (body.SequenceEqual(bytes)) return style;
            return null;
        }

        /// <summary>Records the list style parsed from <paramref name="body"/>; returns the stored copy of the bytes.</summary>
        public byte[] AddListStyle(ReadOnlySpan<byte> body, OdtListStyle style)
        {
            int hash = Hash(body);
            if (!_listStylesByHash.TryGetValue(hash, out List<(byte[] body, OdtListStyle style)>? candidates)) _listStylesByHash[hash] = candidates = new List<(byte[], OdtListStyle)>(1);
            byte[] stored = body.ToArray();
            candidates.Add((stored, style));
            return stored;
        }

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
        public const int Name = 0, Mode = 7, MarginLeft = 8, TextIndent = 9, TabStop = 10, FollowedBy = 11, HasAlignment = 12,
                         StyleName = 13, Href = 14, Width = 15, Height = 16, FontFamily = 17, FontName = 18, LetterSync = 19, SlotCount = 20;
        // xlink:href is matched with the "xlink" prefix every ODF writer uses (the root's own prefix is not passed here)
        private static readonly (bool isStyle, bool isFo, string local)[] Slots =
        [
            (true, false, "name"), (false, false, "level"), (true, false, "num-prefix"), (true, false, "num-suffix"), (true, false, "num-format"),
            (false, false, "start-value"), (false, false, "display-levels"),
            (false, false, "list-level-position-and-space-mode"), (false, true, "margin-left"), (false, true, "text-indent"),
            (false, false, "list-tab-stop-position"), (false, false, "label-followed-by"),
            (false, false, " has-alignment"),   // HasAlignment: no attribute (a name no tag can carry)
            (false, false, "style-name"), (false, false, "xlink:href"), (false, true, "width"), (false, true, "height"),
            (false, true, "font-family"), (true, false, "font-name"), (true, false, "num-letter-sync"),
        ];
        private static ParseNames? _last;

        private readonly string _key;
        public readonly byte[] LevelPrefix, LevelStartMarker, Properties, Alignment, BulletCharacter, TextProperties;
        public readonly byte[][] SlotNames;

        private ParseNames(string key, string text, string style, string fo)
        {
            _key = key;
            LevelPrefix = Encoding.UTF8.GetBytes(text + ":list-level-style-");
            LevelStartMarker = Encoding.UTF8.GetBytes("<" + text + ":list-level-style-");
            Properties = Encoding.UTF8.GetBytes(style + ":list-level-properties");
            Alignment = Encoding.UTF8.GetBytes(style + ":list-level-label-alignment");
            BulletCharacter = Encoding.UTF8.GetBytes(text + ":bullet-char");
            TextProperties = Encoding.UTF8.GetBytes(style + ":text-properties");
            SlotNames = Slots.Select(slot => Encoding.UTF8.GetBytes(slot.local.StartsWith("xlink:", StringComparison.Ordinal) ? slot.local
                                                                     : (slot.isStyle ? style : slot.isFo ? fo : text) + ":" + slot.local)).ToArray();
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
            ("start-value", false) => 5, ("display-levels", false) => 6, ("bullet-char", false) => Name, ("num-letter-sync", true) => LetterSync,
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
                slots[slot] = WordDocumentXml.DecodeAttributeValue(value);
            }
            position += close + 2;
        }
        return position;
    }

    /// <param name="attribute">(local name, is style: namespace else text:) -> value.</param>
    private void AddLevel(string localName, bool isTextNamespace, Func<string, bool, string?> attribute, string layout, LevelExtras extras)
    {
        if (BuildLevel(localName, isTextNamespace, attribute, layout, extras) is { } built) _levels[built.index] = built.level;
    }

    private static readonly XNamespace Xlink = "http://www.w3.org/1999/xlink";

    /// <summary>
    /// A level's label look (xmloff xmlnumi.cxx): text:style-name, the label's character style (:276-278); xlink:href of a
    /// list-level-style-image (:285-287); the bullet font from style:text-properties fo:font-family / style:font-name (:579-583,
    /// :645-694 - a style:font-name is an office:font-face-decls name, resolved to its svg:font-family like a run's font by
    /// <see cref="ResolveFontNames"/>; fo:font-family, read after it there, wins); the image's fo:width / fo:height on
    /// style:list-level-properties (:604-611).
    /// </summary>
    private readonly record struct LevelExtras(string? StyleName, string? ImageHref, string? FontFamily, string? FontName, string? Width, string? Height);

    /// <summary>Separates <see cref="Level.Layout"/>'s geometry from the label's text style name and kept bullet font.</summary>
    public const char LayoutLabelLookSeparator = '\u0001';

    private static (int index, Level level)? BuildLevel(string localName, bool isTextNamespace, Func<string, bool, string?> attribute, string layout, LevelExtras extras)
    {
        if (!int.TryParse(attribute("level", false), NumberStyles.Integer, CultureInfo.InvariantCulture, out int levelNumber)) return null;
        int index = Math.Max(1, Math.Min(MaxLevels, levelNumber)) - 1;
        var level = new Level { Prefix = attribute("num-prefix", true) ?? "", Suffix = attribute("num-suffix", true) ?? "" };
        if (isTextNamespace && localName == "outline-level-style")
        {
            // text:outline-style (ODF 1.3 part 3 16.34): heading numbering; no style:num-format = no number (xmloff xmlnumi.cxx reads an
            // outline level like a number level, its format defaulting to none)
            string? format = attribute("num-format", true);
            level.NumberFormat = string.IsNullOrEmpty(format) ? "" : NumberFormatOf(format, attribute("num-letter-sync", true));
            if (int.TryParse(attribute("start-value", false), NumberStyles.Integer, CultureInfo.InvariantCulture, out int start)) level.StartValue = start < 0 ? 1 : start;
            if (int.TryParse(attribute("display-levels", false), NumberStyles.Integer, CultureInfo.InvariantCulture, out int display)) level.DisplayLevels = Math.Max(1, display);
        }
        else if (isTextNamespace && localName == "list-level-style-number")
        {
            level.NumberFormat = NumberFormatOf(attribute("num-format", true), attribute("num-letter-sync", true));
            if (int.TryParse(attribute("start-value", false), NumberStyles.Integer, CultureInfo.InvariantCulture, out int start)) level.StartValue = start < 0 ? 1 : start;
            if (int.TryParse(attribute("display-levels", false), NumberStyles.Integer, CultureInfo.InvariantCulture, out int display)) level.DisplayLevels = Math.Max(1, display);
        }
        else if (isTextNamespace && localName == "list-level-style-bullet")
        {
            level.IsBullet = true;
            string bullet = attribute("bullet-char", false) ?? "";
            // a style:font-name: its family is known once the font-face-decls are read (ResolveFontNames); until then the name
            if (extras.FontFamily is null && extras.FontName is not null)
                level.Unresolved = new UnresolvedBulletFont(extras.FontName, bullet, layout, extras.StyleName);
            ApplyBullet(level, bullet, extras.FontFamily ?? BulletFont(extras.FontName), layout, extras.StyleName);
            return (index, level);
        }
        else
        {
            // list-level-style-image (ODF 16.37): the package picture its xlink:href names, at its fo:width / fo:height ratio
            // (xmlnumi.cxx:486-509 GraphicBitmap + GraphicSize); embedded office:binary-data or a missing size: a plain bullet
            level.IsBullet = true;
            string? href = extras.ImageHref is { Length: > 0 } link && !link.Contains(':') ? (link.StartsWith("./", StringComparison.Ordinal) ? link.Substring(2) : link) : null;
            double width = extras.Width is null ? 0 : OdtStyleCascade.Twips(extras.Width) ?? 0, height = extras.Height is null ? 0 : OdtStyleCascade.Twips(extras.Height) ?? 0;
            level.BulletCharacter = WordListLabelHtml.PictureLabel(href is not null && WordPictureHtml.ContentTypeOf(href) is not null ? OdtBlockHtmlRenderer.PackagePicturePrefix + href : null,
                                                                   width, height);
        }
        // a number or picture label: only its character style (same layout string as before: an empty bullet-font slot)
        level.Layout = extras.StyleName is null ? layout : layout + LayoutLabelLookSeparator + extras.StyleName + LayoutLabelLookSeparator;
        return (index, level);
    }

    /// <summary>
    /// A bullet level's character and label look. A symbol-font bullet (U+F0xx) as in .docx: Symbol recoded to Unicode,
    /// Wingdings / Webdings keep their code and font, any other font a plain bullet (WordSymbolFontBullet; LibreOffice:
    /// xmlnumi.cxx:456-480 recodes StarBats / StarMath only).
    /// </summary>
    private static void ApplyBullet(Level level, string bullet, string? fontFamily, string layout, string? styleName)
    {
        string? font = fontFamily?.Trim('\'', '"');
        (string resolved, bool keepsFont) = bullet.Length == 0 ? ("\u2022", false) : WordSymbolFontBullet.Resolve(bullet, font);
        level.BulletCharacter = resolved;
        string? bulletFont = keepsFont ? font : null;
        // the stored bullet had a strong left-to-right (symbol-font) code before recoding
        bool isLatinBullet = bullet.Length > 0 && resolved != bullet && !WordListLabelScript.UsesComplexScript(bullet, isParagraphRightToLeft: true);
        level.Layout = styleName is null && bulletFont is null && !isLatinBullet ? layout
            : layout + LayoutLabelLookSeparator + (styleName ?? "") + LayoutLabelLookSeparator + (bulletFont ?? "")
              + (isLatinBullet ? LayoutLabelLookSeparator + "L" : "");
    }

    /// <summary>A bullet level whose style:font-name still has to be looked up in office:font-face-decls.</summary>
    internal sealed record UnresolvedBulletFont(string FontName, string Bullet, string Layout, string? StyleName);

    /// <summary>
    /// Resolves the bullet levels' style:font-name through office:font-face-decls (name -> svg:font-family; a name with no
    /// declaration is taken as the family, as for runs: OdtStyleCascade.FontName). Called by the cascade when the list style
    /// is handed out, after the sections are read: the styles.xml byte filter parses a list style before the XmlReader has
    /// necessarily read the declarations. A level shared by several list styles (ParsedLevels) resolves once.
    /// </summary>
    internal void ResolveFontNames(IReadOnlyDictionary<string, string> fontFamilies)
    {
        if (Volatile.Read(ref _isFontNamesResolved)) return;
        // one lock for all styles: ParsedLevels shares Level objects between list styles, and the index builder and the renderer can ask at once
        lock (FontNameResolution)
        {
            if (_isFontNamesResolved) return;
            EnsureLevelsParsed();                              // the same lock: re-entered
            foreach (Level level in _levels)
            {
                if (level.Unresolved is not { } pending) continue;
                string family = fontFamilies.TryGetValue(pending.FontName, out string? declared) ? declared : pending.FontName;
                ApplyBullet(level, pending.Bullet, BulletFont(family), pending.Layout, pending.StyleName);
                level.Unresolved = null;
            }
            Volatile.Write(ref _isFontNamesResolved, true);
        }
    }

    private static readonly object FontNameResolution = new();
    private bool _isFontNamesResolved;

    private static string LayoutOf(string? mode, bool hasAlignment, string? marginLeft, string? textIndent, string? tabStop, string? followedBy)
    {
        if (mode != "label-alignment" || !hasAlignment) return "";
        int margin = marginLeft is null ? 0 : OdtStyleCascade.Twips(marginLeft) ?? 0;
        int indent = textIndent is null ? 0 : OdtStyleCascade.Twips(textIndent) ?? 0;
        int tab = tabStop is null ? -1 : OdtStyleCascade.Twips(tabStop) ?? -1;
        int follow = followedBy switch { "space" => 1, "nothing" => 2, _ => 0 };
        return string.Join(",", margin, indent, tab, follow);
    }

    public Level LevelAt(int level)
    {
        EnsureLevelsParsed();
        return _levels[Math.Max(0, Math.Min(MaxLevels - 1, level))];
    }

    /// <summary>The label-relevant fields of every level, as the compact index stores the list styles a rescan needs.</summary>
    public void Write(BinaryWriter writer)
    {
        EnsureLevelsParsed();
        foreach (Level level in _levels)
        {
            writer.Write(level.IsBullet);
            writer.Write(level.BulletCharacter);
            writer.Write(level.NumberFormat);
            writer.Write(level.Prefix);
            writer.Write(level.Suffix);
            writer.Write(level.StartValue);
            writer.Write(level.DisplayLevels);
            writer.Write(level.Layout);
        }
    }

    public static OdtListStyle Read(BinaryReader reader)
    {
        var style = new OdtListStyle();
        for (int i = 0; i < MaxLevels; i++)
        {
            style._levels[i] = new Level
            {
                IsBullet = reader.ReadBoolean(), BulletCharacter = reader.ReadString(), NumberFormat = reader.ReadString(),
                Prefix = reader.ReadString(), Suffix = reader.ReadString(), StartValue = reader.ReadInt32(), DisplayLevels = reader.ReadInt32(), Layout = reader.ReadString(),
            };
        }
        return style;
    }

    /// <summary>
    /// The label of an item at <paramref name="level"/> given the current counters of its list:
    /// the last display-levels numbers joined with ".", wrapped in this level's prefix and suffix.
    /// </summary>
    public string FormatLabel(int level, int[] counters) => FormatLabel(level, counters, null);

    /// <summary>True when a number this level shows follows the paragraph's language (<see cref="ListNumberingTypes.DependsOnLanguage"/>).</summary>
    public bool DependsOnLanguage(int level)
    {
        Level current = LevelAt(level);
        if (current.IsBullet) return false;
        for (int i = Math.Max(0, level - current.DisplayLevels + 1); i <= level; i++)
            if (!LevelAt(i).IsBullet && ListNumberingTypes.FromOdfNumberFormat(LevelAt(i).NumberFormat, letterSync: false) is { } type
                && ListNumberingTypes.DependsOnLanguage(type)) return true;
        return false;
    }

    /// <param name="language">The paragraph's language at its first character (BCP 47), null = the default.</param>
    public string FormatLabel(int level, int[] counters, string? language)
    {
        Level current = LevelAt(level);
        // A bullet is drawn alone: LibreOffice's SwBulletPortion gets only the bullet character (sw/source/core/text/txtfld.cxx,
        // SVX_NUM_CHAR_SPECIAL); style:num-prefix / num-suffix on a bullet level (LibreOffice writes the bullet as num-suffix
        // too) are not shown.
        if (current.IsBullet) return current.BulletCharacter;
        var label = new StringBuilder(current.Prefix);
        int first = Math.Max(0, level - current.DisplayLevels + 1);
        for (int i = first; i <= level; i++)
        {
            Level part = LevelAt(i);
            if (part.IsBullet) continue;
            string number = FormatNumber(counters[i] != OdtListNumbering.Unset ? counters[i] : part.StartValue, part.NumberFormat, language);
            if (number.Length == 0) continue;
            if (label.Length > current.Prefix.Length) label.Append('.');
            label.Append(number);
        }
        return label.Append(current.Suffix).ToString();
    }

    /// <summary>
    /// style:num-format as stored: style:num-letter-sync="true" turns "a" / "A" into the repeating letters (aa, bb), kept as
    /// their provider identifiers "aaa" / "AAA" (xmloff xmluconv.cxx:652 convertNumFormat; defaultnumberingprovider.cxx:1093-1094).
    /// </summary>
    public static string NumberFormatOf(string? format, string? letterSync)
        => format is "a" or "A" && letterSync == "true" ? format + format + format : format ?? "1";

    /// <summary>style:num-format -> label text, as LibreOffice draws it (<see cref="ListNumberingTypes"/>); "" = no number.</summary>
    public static string FormatNumber(int value, string format, string? language = null)
        => ListNumberingTypes.FromOdfNumberFormat(format, letterSync: false) is { } type ? ListNumberingTypes.Format(value, type, language) : "";

    /// <summary>A bullet's font name without the OpenType features a LibreOffice font name may carry ("Name:-liga"), which
    /// are not part of the family (see <see cref="OdtStyleCascade.FontWithoutFeatures"/>).</summary>
    private static string? BulletFont(string? font) => font is null ? null : OdtStyleCascade.FontWithoutFeatures(font, out _);
}

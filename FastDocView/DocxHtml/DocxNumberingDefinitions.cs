using System.Text;
using System.Xml.Linq;

namespace FastDocView.DocxHtml;

/// <summary>
/// word/numbering.xml (ECMA-376 §17.9): numId → abstractNumId → per-level number format, start value, label pattern
/// ("%1.", "%1.%2"), restart rule, legal numbering, suffix and label run properties. The level a paragraph uses is looked
/// up the way Open-Xml-PowerTools' ListItemRetriever does it (num/lvlOverride/lvl first — "applied in place of" the
/// abstract level, §17.9.9 — then the abstractNum's lvl); an abstractNum with w:numStyleLink takes its levels from the
/// abstractNum of the num that the linked numbering style's numPr names (§17.9.21). Lists sharing an abstractNum continue
/// one sequence unless a num carries a startOverride. Labels are computed once at index time by
/// <see cref="DocxListCounter"/>, because a windowed renderer never sees the earlier items.
/// </summary>
public sealed class DocxNumberingDefinitions
{
    private static readonly XNamespace W = DocxStyleCascade.W;

    /// <param name="Restart">w:lvlRestart (§17.9.10): null = restart after any higher level; 0 = never; N = when a level 1..N (one-based) is used.</param>
    /// <param name="IsLegal">w:isLgl (§17.9.4): every level shown in this level's text is decimal.</param>
    /// <param name="Suffix">w:suff (§17.9.28): tab (the default), space or nothing after the label.</param>
    /// <param name="RunProperties">The level's w:rPr: the label's own formatting.</param>
    public sealed record LevelDefinition(string Format, string Text, int Start, int? Restart = null, bool IsLegal = false,
                                         string Suffix = "tab", XElement? RunProperties = null,
                                         XElement? ParagraphProperties = null, string Justification = "left", int? PictureBulletId = null,
                                         Html.WordListLabelLegacy? Legacy = null);

    /// <summary>
    /// A w:numPicBullet (17.9.20): the picture's relationship id in numbering.xml's part (VML v:imagedata r:id or a DrawingML
    /// a:blip r:embed) and its stored size, points. Word scales the bullet to the label's font and keeps only this ratio.
    /// </summary>
    public readonly record struct PictureBullet(string RelationshipId, double WidthPt, double HeightPt);

    private Dictionary<int, PictureBullet>? _pictureBullets;
    private readonly byte[]? _partBytes;
    private readonly XElement? _partRoot;

    private readonly Dictionary<int, int> _abstractIdByNumId = new();
    private readonly Dictionary<int, Dictionary<int, int>> _startOverridesByNumId = new();
    private readonly Dictionary<(int numId, int level), LevelDefinition> _overrideLevels = new();
    private readonly Dictionary<(int abstractId, int level), LevelDefinition> _levels = new();
    private readonly Dictionary<int, string> _numStyleLinkByAbstractId = new();
    private readonly Func<string, int>? _numIdOfStyle;

    // Lazy source (numbering.xml cut into its top-level elements): a num / abstractNum is parsed on first use. Null once
    // every definition is in the tables — the part was given parsed, or parsed whole after a fragment failed to parse.
    private DocxNumberingPartSlices? _slices;
    private readonly HashSet<int> _loadedNums = new();
    private readonly HashSet<int> _loadedAbstractNums = new();

    /// <param name="numIdOfStyle">A numbering style's numId (its pPr/numPr), for w:numStyleLink; -1 when it has none.</param>
    public DocxNumberingDefinitions(XDocument? numberingDocument, Func<string, int>? numIdOfStyle = null)
    {
        _numIdOfStyle = numIdOfStyle;
        _partRoot = numberingDocument?.Root;
        ReadWhole(numberingDocument);
    }

    /// <summary>
    /// From the part's bytes: only the lists a window uses are parsed (<see cref="DocxNumberingPartSlices"/>), so a big
    /// document's first window does not pay for a megabyte of numbering it never shows. A part the slicer cannot follow,
    /// or a fragment that does not parse, falls back to the whole-part parse (keeping what precedes the first error).
    /// </summary>
    public DocxNumberingDefinitions(byte[]? numberingPart, Func<string, int>? numIdOfStyle = null)
    {
        _numIdOfStyle = numIdOfStyle;
        if (numberingPart is null) return;
        _partBytes = numberingPart;
        _slices = DocxNumberingPartSlices.TryCut(numberingPart);
        if (_slices is null) ReadWholeBytes(numberingPart);
    }

    private void ReadWholeBytes(byte[] numberingPart)
        => ReadWhole(WordDocumentXml.LoadPartKeepingContentBeforeFirstError(() => new MemoryStream(numberingPart, writable: false)));

    private void ReadWhole(XDocument? numberingDocument)
    {
        XElement? root = numberingDocument?.Root;
        if (root is null) return;
        foreach (XElement num in root.Elements(W + "num")) ReadNum(num);
        foreach (XElement abstractNum in root.Elements(W + "abstractNum")) ReadAbstractNum(abstractNum);
    }

    private void ReadNum(XElement num)
    {
        int? numId = DocxStyleCascade.TwipsAttribute(num, "numId");
        int? abstractId = DocxStyleCascade.TwipsAttribute(num.Element(W + "abstractNumId"), "val");
        if (numId is not { } id || abstractId is not { } abstractNumId) return;
        _abstractIdByNumId[id] = abstractNumId;
        foreach (XElement levelOverride in num.Elements(W + "lvlOverride"))
        {
            int level = DocxStyleCascade.TwipsAttribute(levelOverride, "ilvl") ?? 0;
            if (DocxStyleCascade.TwipsAttribute(levelOverride.Element(W + "startOverride"), "val") is { } start)
            {
                if (!_startOverridesByNumId.TryGetValue(id, out var overrides)) _startOverridesByNumId[id] = overrides = new Dictionary<int, int>();
                overrides[level] = start;
            }
            if (levelOverride.Element(W + "lvl") is { } lvl) _overrideLevels[(id, level)] = ReadLevel(lvl, level);
        }
    }

    private void ReadAbstractNum(XElement abstractNum)
    {
        int? abstractId = DocxStyleCascade.TwipsAttribute(abstractNum, "abstractNumId");
        if (abstractId is not { } id) return;
        if ((string?)abstractNum.Element(W + "numStyleLink")?.Attribute(W + "val") is { } link) _numStyleLinkByAbstractId[id] = link;
        foreach (XElement level in abstractNum.Elements(W + "lvl"))
        {
            int levelIndex = DocxStyleCascade.TwipsAttribute(level, "ilvl") ?? 0;
            _levels[(id, levelIndex)] = ReadLevel(level, levelIndex);
        }
    }

    /// <summary>Parses the w:num elements with this id (lazy source only; every one, in order, as the whole-part parse layers duplicates).</summary>
    private void EnsureNum(int numId)
    {
        if (_slices is null || !_loadedNums.Add(numId) || !_slices.NumsById.TryGetValue(numId, out List<DocxNumberingPartSlices.Slice>? slices)) return;
        foreach (DocxNumberingPartSlices.Slice slice in slices)
        {
            if (ParseSlice(slice) is not { } num) return;       // fell back to the whole part: everything is loaded now
            ReadNum(num);
        }
    }

    private void EnsureAbstractNum(int abstractId)
    {
        if (_slices is null || !_loadedAbstractNums.Add(abstractId) || !_slices.AbstractNumsById.TryGetValue(abstractId, out List<DocxNumberingPartSlices.Slice>? slices)) return;
        foreach (DocxNumberingPartSlices.Slice slice in slices)
        {
            if (ParseSlice(slice) is not { } abstractNum) return;
            ReadAbstractNum(abstractNum);
        }
    }

    /// <summary>One top-level element parsed in the root's namespace context; null after a parse failure, which switched to the whole-part tables.</summary>
    private XElement? ParseSlice(DocxNumberingPartSlices.Slice slice)
    {
        DocxNumberingPartSlices slices = _slices!;
        try
        {
            byte[] closeTag = "</w:numbering>"u8.ToArray();
            using var wrapped = new MemoryStream(slices.RootOpenTag.Length + slice.Length + closeTag.Length);
            wrapped.Write(slices.RootOpenTag, 0, slices.RootOpenTag.Length);
            wrapped.Write(slices.Bytes, slice.Start, slice.Length);
            wrapped.Write(closeTag, 0, closeTag.Length);
            wrapped.Position = 0;
            return WordDocumentXml.ReadFirstChildElement(wrapped, LoadOptions.None, tolerateUndeclaredPrefixes: false, "numbering.xml");
        }
        catch (Exception exception) when (exception is InvalidDataException or System.Xml.XmlException)
        {
            _slices = null;
            _abstractIdByNumId.Clear();
            _startOverridesByNumId.Clear();
            _overrideLevels.Clear();
            _levels.Clear();
            _numStyleLinkByAbstractId.Clear();
            _levelsByList.Clear();
            ReadWholeBytes(slices.Bytes);
            return null;
        }
    }

    private static LevelDefinition ReadLevel(XElement level, int levelIndex)
    {
        // A custom number format is written as mc:AlternateContent (Choice: numFmt val="custom", Fallback: a plain numFmt);
        // the fallback is what every reader without that extension uses.
        XElement? numFmt = level.Element(W + "numFmt")
            ?? level.Elements().Where(e => e.Name.LocalName == "AlternateContent").Elements().Where(e => e.Name.LocalName == "Fallback").Elements(W + "numFmt").FirstOrDefault();
        string format = (string?)numFmt?.Attribute(W + "val") ?? "decimal";
        string text = (string?)level.Element(W + "lvlText")?.Attribute(W + "val") ?? "%" + (levelIndex + 1) + ".";
        int start = DocxStyleCascade.TwipsAttribute(level.Element(W + "start"), "val") ?? 1;
        int? restart = DocxStyleCascade.TwipsAttribute(level.Element(W + "lvlRestart"), "val");
        bool isLegal = level.Element(W + "isLgl") is { } legal && DocxStyleCascade.IsOn(legal.Attribute(W + "val"));
        string suffix = (string?)level.Element(W + "suff")?.Attribute(W + "val") ?? "tab";
        // w:lvlJc (17.9.7): where the label sits against its position; w:pPr (17.9.23): the level's indents and tab stops
        string justification = (string?)level.Element(W + "lvlJc")?.Attribute(W + "val") ?? "left";
        int? pictureBulletId = DocxStyleCascade.TwipsAttribute(level.Element(W + "lvlPicBulletId"), "val");
        // w:legacy (17.9.10): a Word 6 / 95 level - legacySpace / legacyIndent in twips (absent = 0) replace the suffix (WordListLabelLegacy)
        Html.WordListLabelLegacy? legacy = level.Element(W + "legacy") is { } legacyElement && DocxStyleCascade.IsOn(legacyElement.Attribute(W + "legacy"))
            ? new Html.WordListLabelLegacy(DocxStyleCascade.TwipsAttribute(legacyElement, "legacySpace") ?? 0, DocxStyleCascade.TwipsAttribute(legacyElement, "legacyIndent") ?? 0)
            : null;
        return new LevelDefinition(format, text, start, restart, isLegal, suffix, level.Element(W + "rPr"), level.Element(W + "pPr"), justification, pictureBulletId, legacy);
    }

    /// <summary>
    /// The w:numPicBullet with this id, read on first use: the part's leading numPicBullet elements only (the schema puts them
    /// first, CT_Numbering 17.9.17), with a forward-only reader that stops at the first other element. LibreOffice reads the
    /// same picture (NumberingManager.cxx:864-881 numPicBullet / pict, :884-933 lvlPicBulletId: only the aspect ratio kept).
    /// </summary>
    public PictureBullet? PictureBulletOf(int id)
    {
        Dictionary<int, PictureBullet>? bullets = _pictureBullets;
        if (bullets is null)
        {
            // filled locally and published whole: labeller passes may share these definitions
            bullets = new Dictionary<int, PictureBullet>();
            if (_partRoot is not null)
                foreach (XElement bullet in _partRoot.Elements(W + "numPicBullet")) ReadPictureBullet(bullet, bullets);
            else if (_partBytes is not null)
                try { ReadLeadingPictureBullets(_partBytes, bullets); }
                catch (System.Xml.XmlException) { }   // a malformed part: the bullets read before the error stay
            _pictureBullets = bullets;
        }
        return bullets.TryGetValue(id, out PictureBullet bulletFound) ? bulletFound : null;
    }

    private static void ReadLeadingPictureBullets(byte[] part, Dictionary<int, PictureBullet> bullets)
    {
        using var reader = System.Xml.XmlReader.Create(new MemoryStream(part, writable: false),
            new System.Xml.XmlReaderSettings { IgnoreWhitespace = true, IgnoreComments = true, DtdProcessing = System.Xml.DtdProcessing.Ignore });
        if (!reader.ReadToFollowing("numbering", W.NamespaceName) || reader.IsEmptyElement) return;
        reader.Read();
        while (!reader.EOF && reader.NodeType != System.Xml.XmlNodeType.EndElement)
        {
            if (reader.NodeType != System.Xml.XmlNodeType.Element) { reader.Read(); continue; }
            if (reader.LocalName != "numPicBullet" || reader.NamespaceURI != W.NamespaceName) return;
            ReadPictureBullet((XElement)XNode.ReadFrom(reader), bullets);
        }
    }

    private static readonly XNamespace RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static void ReadPictureBullet(XElement bullet, Dictionary<int, PictureBullet> bullets)
    {
        if (DocxStyleCascade.TwipsAttribute(bullet, "numPicBulletId") is not { } id) return;
        // VML (what Word writes): v:shape style width / height + v:imagedata r:id; DrawingML (w:drawing): wp:extent cx / cy + a:blip r:embed
        string? relationshipId = null;
        double width = 0, height = 0;
        foreach (XElement element in bullet.Descendants())
        {
            switch (element.Name.LocalName)
            {
                case "imagedata": relationshipId ??= (string?)element.Attribute(RelationshipsNamespace + "id"); break;
                case "blip": relationshipId ??= (string?)element.Attribute(RelationshipsNamespace + "embed"); break;
                case "shape" when (string?)element.Attribute("style") is { } style:
                    foreach (string declaration in style.Split(';'))
                    {
                        int colon = declaration.IndexOf(':');
                        if (colon < 0) continue;
                        string name = declaration.Substring(0, colon).Trim();
                        if (name == "width") width = DocxBlockHtmlRenderer.VmlLengthPoints(declaration.Substring(colon + 1).Trim());
                        else if (name == "height") height = DocxBlockHtmlRenderer.VmlLengthPoints(declaration.Substring(colon + 1).Trim());
                    }
                    break;
                case "extent" when width <= 0:
                    if (long.TryParse((string?)element.Attribute("cx"), out long cx) && long.TryParse((string?)element.Attribute("cy"), out long cy))
                        (width, height) = (cx / 12700.0, cy / 12700.0);
                    break;
            }
        }
        if (relationshipId is not null) bullets[id] = new PictureBullet(relationshipId, width, height);
    }

    /// <summary>No list at all (with the lazy source: no w:num with an id — a num without an abstractNumId then resolves to no list, as before).</summary>
    public bool IsEmpty => _slices is { } slices ? slices.NumsById.Count == 0 : _abstractIdByNumId.Count == 0;

    /// <summary>The abstractNum whose levels and counters a num uses — through w:numStyleLink when the abstractNum is only a link.</summary>
    public int? AbstractIdOf(int numId)
    {
        EnsureNum(numId);
        if (!_abstractIdByNumId.TryGetValue(numId, out int abstractId)) return null;
        EnsureAbstractNum(abstractId);
        if (!_numStyleLinkByAbstractId.ContainsKey(abstractId)) return abstractId;      // the common case: no link, nothing to allocate
        var seen = new HashSet<int> { abstractId };
        while (_numStyleLinkByAbstractId.TryGetValue(abstractId, out string? style))
        {
            int linkedNumId = _numIdOfStyle?.Invoke(style) ?? -1;
            EnsureNum(linkedNumId);
            if (!_abstractIdByNumId.TryGetValue(linkedNumId, out int linkedAbstractId) || !seen.Add(linkedAbstractId)) break;
            abstractId = linkedAbstractId;
            EnsureAbstractNum(abstractId);
        }
        return abstractId;
    }

    /// <summary>The level definition a paragraph of <paramref name="numId"/> uses: the num's lvlOverride/lvl, else the abstract level.</summary>
    public LevelDefinition? LevelOf(int numId, int abstractId, int level)
    {
        EnsureNum(numId);
        EnsureAbstractNum(abstractId);
        return _overrideLevels.TryGetValue((numId, level), out LevelDefinition? own) ? own
             : _levels.TryGetValue((abstractId, level), out LevelDefinition? definition) ? definition : null;
    }

    private readonly Dictionary<(int numId, int abstractId), LevelDefinition?[]> _levelsByList = new();

    /// <summary>All nine levels a num uses (<see cref="LevelOf"/> per level), resolved once per list: the counter asks for them on every list paragraph.</summary>
    public LevelDefinition?[] LevelsOf(int numId, int abstractId)
    {
        if (_levelsByList.TryGetValue((numId, abstractId), out LevelDefinition?[]? levels)) return levels;
        levels = new LevelDefinition?[DocxListCounter.MaxLevels];
        for (int level = 0; level < levels.Length; level++) levels[level] = LevelOf(numId, abstractId, level);
        _levelsByList[(numId, abstractId)] = levels;
        return levels;
    }

    public int? StartOverrideOf(int numId, int level)
    {
        EnsureNum(numId);
        return _startOverridesByNumId.TryGetValue(numId, out var overrides) && overrides.TryGetValue(level, out int start) ? start : null;
    }

    /// <summary>Fills the "%1.%2" pattern of a level with the current counters of all levels; with isLgl every number is decimal.</summary>
    public string FormatLabel(int numId, int abstractId, int level, int[] counters)
    {
        LevelDefinition?[] levels = LevelsOf(numId, abstractId);
        if (levels[level] is not { } definition) return "•";
        if (definition.Format == "bullet") return BulletCharacter(definition.Text);
        if (definition.Format == "none") return "";
        var label = new StringBuilder(definition.Text.Length + 8);
        for (int i = 0; i < definition.Text.Length; i++)
        {
            char c = definition.Text[i];
            if (c == '%' && i + 1 < definition.Text.Length && definition.Text[i + 1] is >= '1' and <= '9')
            {
                int placeholderLevel = definition.Text[i + 1] - '1';
                string format = definition.IsLegal ? "decimal"
                              : placeholderLevel == level ? definition.Format : levels[placeholderLevel]?.Format ?? "decimal";
                label.Append(FormatNumber(counters[placeholderLevel], format));
                i++;
            }
            else label.Append(c);
        }
        return label.ToString();
    }

    /// <summary>A bullet's level text as written; a symbol-font code is resolved against the level's font by
    /// <see cref="Html.WordSymbolFontBullet"/> where the label's run properties are known.</summary>
    private static string BulletCharacter(string text) => text.Length == 0 ? "•" : text;

    /// <summary>The font a level's w:rPr/w:rFonts names for its label (ascii, else hAnsi, else cs).</summary>
    public static string? FontOf(XElement? runProperties)
    {
        XElement? fonts = runProperties?.Element(W + "rFonts");
        return (string?)fonts?.Attribute(W + "ascii") ?? (string?)fonts?.Attribute(W + "hAnsi");
    }

    /// <summary>
    /// A counter in an ECMA-376 ST_NumberFormat (17.18.59). Every result was compared with Word 16's own ListFormat.ListString
    /// (test-data/word-lists, ListNumberFormatsFidelityTest); formats not listed here (CJK / Korean / Thai / Vietnamese
    /// counting words, kana orders) show decimal, as LibreOffice does for numbering types it lacks.
    /// </summary>
    public static string FormatNumber(int value, string format) => format switch
    {
        "lowerLetter" => Letters(value, 'a'),
        "upperLetter" => Letters(value, 'A'),
        "lowerRoman" => Roman(value).ToLowerInvariant(),
        "upperRoman" => Roman(value),
        "hebrew1" => HebrewNumeral(value),
        "hebrew2" => HebrewLetters(value),
        "decimalZero" => value.ToString("00"),
        "none" => "",
        "ordinal" => value.ToString() + EnglishOrdinalSuffix(value),
        "cardinalText" => value < 0 ? value.ToString() : Capitalize(Fields.WordFieldFormat.CardinalText(value)),
        "ordinalText" => value < 0 ? value.ToString() : Capitalize(Fields.WordFieldFormat.OrdinalText(value)),
        "decimalFullWidth" or "decimalFullWidth2" => Digits(value, (char)0xFF10),
        "hindiNumbers" => Digits(value, (char)0x0966),
        "thaiNumbers" => Digits(value, (char)0x0E50),
        "ideographDigital" or "japaneseDigitalTenThousand" => Digits(value, '\0', IdeographDigits),
        "taiwaneseDigital" => Digits(value, '\0', (char)0x25CB + IdeographDigits.Substring(1)),     // Word 16: 10 = one + white circle
        "koreanDigital2" => Digits(value, '\0', (char)0x96F6 + IdeographDigits.Substring(1)),       // Word 16: 10 = one + U+96F6
        "decimalEnclosedCircle" or "decimalEnclosedCircleChinese" => value is >= 1 and <= 20 ? ((char)(0x2460 + value - 1)).ToString() : value.ToString(),
        "decimalEnclosedFullstop" => value is >= 1 and <= 20 ? ((char)(0x2488 + value - 1)).ToString() : value.ToString(),
        "decimalEnclosedParen" => value is >= 1 and <= 20 ? ((char)(0x2474 + value - 1)).ToString() : value.ToString(),
        "ideographEnclosedCircle" => value is >= 1 and <= 10 ? ((char)(0x3220 + value - 1)).ToString() : value.ToString(),
        "ideographTraditional" => Cycle(value, HeavenlyStems),
        "ideographZodiac" => Cycle(value, EarthlyBranches),
        "ideographZodiacTraditional" => value <= 0 ? value.ToString()
            : HeavenlyStems[(value - 1) % 10].ToString() + EarthlyBranches[(value - 1) % 12],
        "numberInDash" => "- " + value + " -",
        "russianLower" => RepeatedLetters(value, RussianLetters, toUpper: false),
        "russianUpper" => RepeatedLetters(value, RussianLetters, toUpper: true),
        "thaiLetters" => RepeatedLetters(value, ThaiLetters, toUpper: false),
        "arabicAlpha" => RepeatedLetters(value, ArabicAlphaLetters, toUpper: false, joinerAfter: (char)0x200C),
        "arabicAbjad" => RepeatedLetters(value, ArabicAbjadLetters, toUpper: false, joinerBefore: (char)0x200C),
        _ => value.ToString(),
    };

    // Word 16 (ListString): 10 = U+4E00 U+3007, 1234 = one two three four - each decimal digit as an ideograph
    private static readonly string IdeographDigits = "〇一二三四五六七八九";
    private static readonly string HeavenlyStems = "甲乙丙丁戊己庚辛壬癸";
    private static readonly string EarthlyBranches = "子丑寅卯辰巳午未申酉戍亥";
    // Word's Russian list letters: 29 (no yo, short i, hard / soft sign): 10 = ka, 27 = e, 50 = kha kha
    private static readonly string RussianLetters = "абвгдежзиклмнопрстуфхцчшщыэюя";
    // LibreOffice defaultnumberingprovider.cxx:103 table_Alphabet_th less the two vowels U+0E24 / U+0E26: Word repeats after 41 (50 = 9th letter twice)
    private static readonly string ThaiLetters = "กขคงจฉชซฌญฎฏฐฑฒณดตถทธนบปผฝพฟภมยรลวศษสหฬอฮ";
    // defaultnumberingprovider.cxx:89 table_Alphabet_ar / :96 table_Alphabet_ar_abjad (Word repeats the letter; ZWNJ after / before each, as its ListString)
    private static readonly string ArabicAlphaLetters = "أبتثجحخدذرزسشصضطظعغفقكلمنهوي";
    private static readonly string ArabicAbjadLetters = "أبجدهوزحطيكلمنسعفصقرشتثخذضظغ";

    private static string Digits(int value, char zero, string? digits = null)
    {
        string plain = value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        var result = new StringBuilder(plain.Length);
        foreach (char c in plain) result.Append(c is >= '0' and <= '9' ? (digits is null ? (char)(zero + (c - '0')) : digits[c - '0']) : c);
        return result.ToString();
    }

    private static string Cycle(int value, string symbols) => value <= 0 ? value.ToString() : symbols[(value - 1) % symbols.Length].ToString();

    /// <summary>Word's letter lists: the letter of (n - 1) mod count, written (n - 1) / count + 1 times.</summary>
    private static string RepeatedLetters(int value, string letters, bool toUpper, char joinerAfter = '\0', char joinerBefore = '\0')
    {
        if (value <= 0) return "";
        char letter = letters[(value - 1) % letters.Length];
        if (toUpper) letter = char.ToUpperInvariant(letter);
        int repeat = (value - 1) / letters.Length + 1;
        var result = new StringBuilder(repeat * 2);
        for (int i = 0; i < repeat; i++)
        {
            if (joinerBefore != '\0') result.Append(joinerBefore);
            result.Append(letter);
            if (joinerAfter != '\0') result.Append(joinerAfter);
        }
        return result.ToString();
    }

    private static string EnglishOrdinalSuffix(int value)
    {
        int lastTwo = Math.Abs(value) % 100;
        if (lastTwo is >= 11 and <= 13) return "th";
        return (Math.Abs(value) % 10) switch { 1 => "st", 2 => "nd", 3 => "rd", _ => "th" };
    }

    private static string Capitalize(string text) => text.Length == 0 ? text : char.ToUpperInvariant(text[0]) + text.Substring(1);

    private static string Letters(int value, char first)
    {
        if (value <= 0) return "";
        int repeat = (value - 1) / 26 + 1;
        return new string((char)(first + (value - 1) % 26), repeat);
    }

    private static string Roman(int value)
    {
        if (value <= 0 || value >= 4000) return value.ToString();
        (int number, string numeral)[] table = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")];
        var roman = new StringBuilder();
        foreach ((int number, string numeral) in table) while (value >= number) { roman.Append(numeral); value -= number; }
        return roman.ToString();
    }

    private static readonly string[] HebrewOnes = ["", "א", "ב", "ג", "ד", "ה", "ו", "ז", "ח", "ט"];
    private static readonly string[] HebrewTens = ["", "י", "כ", "ל", "מ", "נ", "ס", "ע", "פ", "צ"];
    private static readonly string[] HebrewHundreds = ["", "ק", "ר", "ש", "ת", "תק", "תר", "תש", "תת", "תתק"];

    /// <summary>
    /// Word's hebrew1: gematria letters (טו / טז for 15 / 16) with no geresh or gershayim, cycling after 392 — what
    /// Word 16 shows (ListFormat.ListString of synthetic lists, 2026-10-28: 11 → יא, 15 → טו, 115 → קטו, 500 → קח,
    /// 1000 → רטז, 5784 → רצו).
    /// </summary>
    public static string HebrewNumeral(int value)
    {
        if (value <= 0) return value.ToString();
        value = (value - 1) % 392 + 1;
        var letters = new StringBuilder();
        letters.Append(HebrewHundreds[value / 100]);
        int tensAndOnes = value % 100;
        if (tensAndOnes == 15) letters.Append("טו");
        else if (tensAndOnes == 16) letters.Append("טז");
        else letters.Append(HebrewTens[tensAndOnes / 10]).Append(HebrewOnes[tensAndOnes % 10]);
        return letters.ToString();
    }

    /// <summary>Word's hebrew2 list format: one TAV per full 22, then the letter (Word 16 ListString: 27 = TAV HE, 50 = TAV TAV VAV, 100 = four TAV + LAMED).</summary>
    public static string HebrewLetters(int value)
    {
        if (value <= 0) return "";
        const string alphabet = "אבגדהוזחטיכלמנסעפצקרשת";
        return new string(alphabet[21], (value - 1) / alphabet.Length) + alphabet[(value - 1) % alphabet.Length];
    }

}

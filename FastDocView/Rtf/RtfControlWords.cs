namespace FastDocView.Rtf;

/// <summary>The control words the RTF reader acts on. Anything else is either a destination it skips or a word it ignores.</summary>
internal enum RtfWord : byte
{
    Unknown = 0,
    /// <summary>A destination the reader does not render: the whole group is skipped (header, info, pict, bookmarks, ...).</summary>
    SkippedDestination,
    // ── document / header ──
    Rtf, Ansi, Mac, Pc, Pca, AnsiCodePage, DefaultFont, FontTable, FontCharset, FontCodePage, Font, AssociatedFont,
    ColorTable, Red, Green, Blue, StyleSheet, Style, CharacterStyle, SectionStyle, TableStyle, OutlineLevel,
    ListTable, List, ListId, ListLevel, LevelNumberFormat, LevelNumberFormatNew, LevelStartAt, LevelText, LevelNumbers, LevelNoRestart, LevelLegal, LevelJustification, LevelFollow,
    ListOverrideTable, ListOverride, ListOverrideLevel, ListOverrideStartAt, ListOverrideIndex, ListLevelIndex,
    FootnoteNumberingArabic, FootnoteNumberingLowerAlpha, FootnoteNumberingUpperAlpha, FootnoteNumberingLowerRoman, FootnoteNumberingUpperRoman, FootnoteNumberingChicago,
    EndnoteNumberingArabic, EndnoteNumberingLowerAlpha, EndnoteNumberingUpperAlpha, EndnoteNumberingLowerRoman, EndnoteNumberingUpperRoman, EndnoteNumberingChicago,
    FootnoteStart, EndnoteStart, RightToLeftDocument,
    //  page setup (document defaults, then per section) 
    PaperWidth, LeftMargin, RightMargin, Gutter, SectionDefault, SectionPageWidth, SectionLeftMargin, SectionRightMargin, SectionGutter,
    // ── destinations rendered ──
    Footnote, FootnoteIsEndnote, Field, FieldResult, ListText, ParagraphNumberText, UnicodePair, UnicodeDestination,
    Picture, ShapePicture, NonShapePicture, Shape, ObjectGroup, MathGroup, NestedTableProperties,
    // ── unicode / special characters ──
    UnicodeSkipCount, UnicodeCharacter, Binary, IgnorableDestination, HexCharacter,
    Paragraph, Section, Page, Line, Tab, Cell, Row, NestedCell, NestedRow, Column,
    EmDash, EnDash, Bullet, LeftQuote, RightQuote, LeftDoubleQuote, RightDoubleQuote, EmSpace, EnSpace, QuarterEmSpace,
    ZeroWidthJoiner, ZeroWidthNonJoiner, ZeroWidthBreakOpportunity, ZeroWidthNoBreak, LeftToRightMark, RightToLeftMark,
    NonBreakingSpace, OptionalHyphen, NonBreakingHyphen, Backslash, OpenBrace, CloseBrace, FootnoteMark, SubEntry, Formula,
    // ── paragraph ──
    ParagraphDefault, AlignLeft, AlignRight, AlignCenter, AlignJustify, AlignDistribute,
    LeftIndent, RightIndent, StartIndent, EndIndent, FirstLineIndent, SpaceBefore, SpaceAfter, LineSpacing, LineSpacingMultiple,
    RightToLeftParagraph, LeftToRightParagraph, InTable, NestingLevel, PageBreakBefore,
    // ── table row ──
    RowDefault, CellBoundary, CellShading, RightToLeftRow, LeftToRightRow,
    // ── character ──
    Plain, Bold, Italic, AssociatedBold, AssociatedItalic, Underline, UnderlineDouble, UnderlineDotted, UnderlineDashed, UnderlineWave,
    UnderlineThick, UnderlineWords, UnderlineNone, Strike, StrikeDouble, FontSize, AssociatedFontSize, ForegroundColor, BackgroundColor,
    CharacterShading, Highlight, Superscript, Subscript, NoSuperSubscript, RaisedPosition, LoweredPosition, Hidden, AllCaps, SmallCaps,
    Outline, Shadow, Emboss, Imprint, CharacterSpacing, RightToLeftCharacters, LeftToRightCharacters,
}

/// <summary>
/// Control-word lookup without allocating a string per token: names are hashed as ASCII bytes.
/// The skipped destinations are LibreOffice's RTFControlType::DESTINATION set
/// (sw/source/writerfilter/rtftok/rtftokenizer.cxx, master 2026-10) minus the ones this reader renders.
/// </summary>
internal static class RtfControlWords
{
    private const int TableSize = 2048;
    private static readonly byte[]?[] Names = new byte[TableSize][];
    private static readonly RtfWord[] Words = new RtfWord[TableSize];

    /// <summary>LibreOffice rtftokenizer.cxx: every control word of type DESTINATION.</summary>
    internal static readonly string[] LibreOfficeDestinations =
    [
        "aftncn", "aftnsep", "aftnsepc", "annotation", "atnauthor", "atndate", "atnicn", "atnid", "atnparent", "atnref", "atntime", "atrfend", "atrfstart",
        "author", "background", "bkmkend", "bkmkstart", "blipuid", "buptim", "category", "colorschememapping", "colortbl", "comment", "company", "creatim",
        "datafield", "datastore", "defchp", "defpap", "do", "doccomm", "docvar", "dptxbxtext", "ebcend", "ebcstart", "factoidname", "falt", "fchars",
        "ffdeftext", "ffentrymcr", "ffexitmcr", "ffformat", "ffhelptext", "ffl", "ffname", "ffstattext", "field", "file", "filetbl", "fldinst", "fldrslt",
        "fldtype", "fname", "fontemb", "fontfile", "fonttbl", "footer", "footerf", "footerl", "footerr", "footnote", "formfield", "ftncn", "ftnsep",
        "ftnsepc", "g", "generator", "gridtbl", "header", "headerf", "headerl", "headerr", "hl", "hlfr", "hlinkbase", "hlloc", "hlsrc", "hsv", "htmltag",
        "info", "keycode", "keywords", "latentstyles", "lchars", "levelnumbers", "leveltext", "lfolevel", "linkval", "list", "listlevel", "listname",
        "listoverride", "listoverridetable", "listpicture", "liststylename", "listtable", "listtext", "lsdlockedexcept", "macc", "maccPr", "mailmerge",
        "maln", "malnScr", "manager", "margPr", "mbar", "mbarPr", "mbaseJc", "mbegChr", "mborderBox", "mborderBoxPr", "mbox", "mboxPr", "mchr", "mcount",
        "mctrlPr", "md", "mdeg", "mdegHide", "mden", "mdiff", "mdPr", "me", "mendChr", "meqArr", "meqArrPr", "mf", "mfName", "mfPr", "mfunc", "mfuncPr",
        "mgroupChr", "mgroupChrPr", "mgrow", "mhideBot", "mhideLeft", "mhideRight", "mhideTop", "mhtmltag", "mlim", "mlimloc", "mlimLoc", "mlimlow",
        "mlimLow", "mlimlowPr", "mlimLowPr", "mlimupp", "mlimUpp", "mlimuppPr", "mlimUppPr", "mm", "mmaddfieldname", "mmath", "mmathPict", "mmathPr",
        "mmaxdist", "mmc", "mmcJc", "mmconnectstr", "mmconnectstrdata", "mmcPr", "mmcs", "mmdatasource", "mmheadersource", "mmmailsubject", "mmodso",
        "mmodsofilter", "mmodsofldmpdata", "mmodsomappedname", "mmodsoname", "mmodsorecipdata", "mmodsosort", "mmodsosrc", "mmodsotable", "mmodsoudl",
        "mmodsoudldata", "mmodsouniquetag", "mmPr", "mmquery", "mmr", "mnary", "mnaryPr", "mnoBreak", "mnum", "mobjDist", "moMath", "moMathPara",
        "moMathParaPr", "mopEmu", "mphant", "mphantPr", "mplcHide", "mpos", "mr", "mrad", "mradPr", "mrPr", "msepChr", "mshow", "mshp", "msPre", "msPrePr",
        "msSub", "msSubPr", "msSubSup", "msSubSupPr", "msSup", "msSupPr", "mstrikeBLTR", "mstrikeH", "mstrikeTLBR", "mstrikeV", "msub", "msubHide", "msup",
        "msupHide", "mtransp", "mtype", "mvertJc", "mvfmf", "mvfml", "mvtof", "mvtol", "mzeroAsc", "mzeroDesc", "mzeroWid", "nesttableprops", "nextfile",
        "nonesttables", "objalias", "objclass", "objdata", "object", "objname", "objsect", "objtime", "oldcprops", "oldpprops", "oldsprops", "oldtprops",
        "oleclsid", "operator", "panose", "password", "passwordhash", "pgp", "pgptbl", "picprop", "pict", "pn", "pnseclvl", "pntext", "pntxta", "pntxtb",
        "printim", "private", "propname", "protend", "protstart", "protusertbl", "pxe", "result", "revtbl", "revtim", "rsidtbl", "rtf", "rxe", "shp",
        "shpgrp", "shpinst", "shppict", "shprslt", "shptxt", "sn", "sp", "staticval", "stylesheet", "subject", "sv", "svb", "tc", "template", "themedata",
        "title", "txe", "ud", "upr", "userprops", "wgrffmtfilter", "windowcaption", "writereservation", "writereservhash", "xe", "xform", "xmlattrname",
        "xmlattrvalue", "xmlclose", "xmlname", "xmlnstbl", "xmlopen", "flymaincnt",
    ];

    /// <summary>The words acted on, with their meaning (RTF 1.5 spec + LibreOffice rtfdispatch*.cxx).</summary>
    internal static readonly (string name, RtfWord word)[] HandledWords =
    [
        ("rtf", RtfWord.Rtf), ("ansi", RtfWord.Ansi), ("mac", RtfWord.Mac), ("pc", RtfWord.Pc), ("pca", RtfWord.Pca), ("ansicpg", RtfWord.AnsiCodePage),
        ("deff", RtfWord.DefaultFont), ("fonttbl", RtfWord.FontTable), ("fcharset", RtfWord.FontCharset), ("cpg", RtfWord.FontCodePage), ("f", RtfWord.Font),
        ("af", RtfWord.AssociatedFont), ("colortbl", RtfWord.ColorTable), ("red", RtfWord.Red), ("green", RtfWord.Green), ("blue", RtfWord.Blue),
        ("stylesheet", RtfWord.StyleSheet), ("s", RtfWord.Style), ("cs", RtfWord.CharacterStyle), ("ds", RtfWord.SectionStyle), ("ts", RtfWord.TableStyle),
        ("outlinelevel", RtfWord.OutlineLevel), ("listtable", RtfWord.ListTable), ("list", RtfWord.List), ("listid", RtfWord.ListId),
        ("listlevel", RtfWord.ListLevel), ("levelnfc", RtfWord.LevelNumberFormat), ("levelnfcn", RtfWord.LevelNumberFormatNew),
        ("levelstartat", RtfWord.LevelStartAt), ("leveltext", RtfWord.LevelText), ("levelnumbers", RtfWord.LevelNumbers),
        ("levelnorestart", RtfWord.LevelNoRestart), ("levellegal", RtfWord.LevelLegal),
        ("leveljc", RtfWord.LevelJustification), ("levelfollow", RtfWord.LevelFollow), ("listoverridetable", RtfWord.ListOverrideTable),
        ("listoverride", RtfWord.ListOverride), ("lfolevel", RtfWord.ListOverrideLevel), ("listoverridestartat", RtfWord.ListOverrideStartAt),
        ("ls", RtfWord.ListOverrideIndex), ("ilvl", RtfWord.ListLevelIndex),
        ("ftnnar", RtfWord.FootnoteNumberingArabic), ("ftnnalc", RtfWord.FootnoteNumberingLowerAlpha), ("ftnnauc", RtfWord.FootnoteNumberingUpperAlpha),
        ("ftnnrlc", RtfWord.FootnoteNumberingLowerRoman), ("ftnnruc", RtfWord.FootnoteNumberingUpperRoman), ("ftnnchi", RtfWord.FootnoteNumberingChicago),
        ("aftnnar", RtfWord.EndnoteNumberingArabic), ("aftnnalc", RtfWord.EndnoteNumberingLowerAlpha), ("aftnnauc", RtfWord.EndnoteNumberingUpperAlpha),
        ("aftnnrlc", RtfWord.EndnoteNumberingLowerRoman), ("aftnnruc", RtfWord.EndnoteNumberingUpperRoman), ("aftnnchi", RtfWord.EndnoteNumberingChicago),
        ("ftnstart", RtfWord.FootnoteStart), ("aftnstart", RtfWord.EndnoteStart), ("rtldoc", RtfWord.RightToLeftDocument),
        ("paperw", RtfWord.PaperWidth), ("margl", RtfWord.LeftMargin), ("margr", RtfWord.RightMargin), ("gutter", RtfWord.Gutter),
        ("sectd", RtfWord.SectionDefault), ("pgwsxn", RtfWord.SectionPageWidth), ("marglsxn", RtfWord.SectionLeftMargin),
        ("margrsxn", RtfWord.SectionRightMargin), ("guttersxn", RtfWord.SectionGutter),
        ("footnote", RtfWord.Footnote), ("ftnalt", RtfWord.FootnoteIsEndnote), ("field", RtfWord.Field), ("fldrslt", RtfWord.FieldResult),
        ("listtext", RtfWord.ListText), ("pntext", RtfWord.ParagraphNumberText), ("upr", RtfWord.UnicodePair), ("ud", RtfWord.UnicodeDestination),
        ("pict", RtfWord.Picture), ("shppict", RtfWord.ShapePicture), ("nonshppict", RtfWord.NonShapePicture), ("shp", RtfWord.Shape),
        ("object", RtfWord.ObjectGroup), ("nesttableprops", RtfWord.NestedTableProperties), ("mmath", RtfWord.MathGroup), ("moMath", RtfWord.MathGroup), ("moMathPara", RtfWord.MathGroup),
        ("uc", RtfWord.UnicodeSkipCount), ("u", RtfWord.UnicodeCharacter), ("bin", RtfWord.Binary),
        ("par", RtfWord.Paragraph), ("sect", RtfWord.Section), ("page", RtfWord.Page), ("line", RtfWord.Line), ("tab", RtfWord.Tab),
        ("cell", RtfWord.Cell), ("row", RtfWord.Row), ("nestcell", RtfWord.NestedCell), ("nestrow", RtfWord.NestedRow), ("column", RtfWord.Column),
        ("emdash", RtfWord.EmDash), ("endash", RtfWord.EnDash), ("bullet", RtfWord.Bullet), ("lquote", RtfWord.LeftQuote), ("rquote", RtfWord.RightQuote),
        ("ldblquote", RtfWord.LeftDoubleQuote), ("rdblquote", RtfWord.RightDoubleQuote), ("emspace", RtfWord.EmSpace), ("enspace", RtfWord.EnSpace),
        ("qmspace", RtfWord.QuarterEmSpace), ("zwj", RtfWord.ZeroWidthJoiner), ("zwnj", RtfWord.ZeroWidthNonJoiner), ("zwbo", RtfWord.ZeroWidthBreakOpportunity),
        ("zwnbo", RtfWord.ZeroWidthNoBreak), ("ltrmark", RtfWord.LeftToRightMark), ("rtlmark", RtfWord.RightToLeftMark), ("chftn", RtfWord.FootnoteMark),
        ("pard", RtfWord.ParagraphDefault), ("ql", RtfWord.AlignLeft), ("qr", RtfWord.AlignRight), ("qc", RtfWord.AlignCenter), ("qj", RtfWord.AlignJustify),
        ("qd", RtfWord.AlignDistribute), ("li", RtfWord.LeftIndent), ("ri", RtfWord.RightIndent), ("lin", RtfWord.StartIndent), ("rin", RtfWord.EndIndent),
        ("fi", RtfWord.FirstLineIndent), ("sb", RtfWord.SpaceBefore), ("sa", RtfWord.SpaceAfter), ("sl", RtfWord.LineSpacing), ("slmult", RtfWord.LineSpacingMultiple),
        ("rtlpar", RtfWord.RightToLeftParagraph), ("ltrpar", RtfWord.LeftToRightParagraph), ("intbl", RtfWord.InTable), ("itap", RtfWord.NestingLevel),
        ("pagebb", RtfWord.PageBreakBefore),
        ("trowd", RtfWord.RowDefault), ("cellx", RtfWord.CellBoundary), ("clcbpat", RtfWord.CellShading), ("rtlrow", RtfWord.RightToLeftRow), ("ltrrow", RtfWord.LeftToRightRow),
        ("plain", RtfWord.Plain), ("b", RtfWord.Bold), ("i", RtfWord.Italic), ("ab", RtfWord.AssociatedBold), ("ai", RtfWord.AssociatedItalic),
        ("ul", RtfWord.Underline), ("uldb", RtfWord.UnderlineDouble), ("uld", RtfWord.UnderlineDotted), ("ulthd", RtfWord.UnderlineDotted),
        ("uldash", RtfWord.UnderlineDashed), ("uldashd", RtfWord.UnderlineDashed), ("uldashdd", RtfWord.UnderlineDashed), ("ulldash", RtfWord.UnderlineDashed),
        ("ulthdash", RtfWord.UnderlineDashed), ("ulthdashd", RtfWord.UnderlineDashed), ("ulthdashdd", RtfWord.UnderlineDashed), ("ulthldash", RtfWord.UnderlineDashed),
        ("ulwave", RtfWord.UnderlineWave), ("ulhwave", RtfWord.UnderlineWave), ("ululdbwave", RtfWord.UnderlineWave), ("ulth", RtfWord.UnderlineThick),
        ("ulw", RtfWord.UnderlineWords), ("ulnone", RtfWord.UnderlineNone), ("strike", RtfWord.Strike), ("striked", RtfWord.StrikeDouble),
        ("fs", RtfWord.FontSize), ("afs", RtfWord.AssociatedFontSize), ("cf", RtfWord.ForegroundColor), ("cb", RtfWord.BackgroundColor),
        ("chcbpat", RtfWord.CharacterShading), ("highlight", RtfWord.Highlight), ("super", RtfWord.Superscript), ("sub", RtfWord.Subscript),
        ("nosupersub", RtfWord.NoSuperSubscript), ("up", RtfWord.RaisedPosition), ("dn", RtfWord.LoweredPosition), ("v", RtfWord.Hidden),
        ("caps", RtfWord.AllCaps), ("scaps", RtfWord.SmallCaps), ("outl", RtfWord.Outline), ("shad", RtfWord.Shadow), ("embo", RtfWord.Emboss),
        ("impr", RtfWord.Imprint), ("expndtw", RtfWord.CharacterSpacing), ("rtlch", RtfWord.RightToLeftCharacters), ("ltrch", RtfWord.LeftToRightCharacters),
    ];

    static RtfControlWords()
    {
        foreach (string name in LibreOfficeDestinations) Add(name, RtfWord.SkippedDestination);
        foreach ((string name, RtfWord word) in HandledWords) Add(name, word);
    }

    private static void Add(string name, RtfWord word)
    {
        byte[] bytes = new byte[name.Length];
        for (int i = 0; i < name.Length; i++) bytes[i] = (byte)name[i];
        int slot = Hash(bytes) & (TableSize - 1);
        while (Names[slot] is { } existing && !existing.AsSpan().SequenceEqual(bytes)) slot = (slot + 1) & (TableSize - 1);
        Names[slot] = bytes;
        Words[slot] = word;
    }

    private static int Hash(ReadOnlySpan<byte> name)
    {
        uint hash = 2166136261;
        foreach (byte b in name) hash = (hash ^ b) * 16777619;
        return (int)(hash ^ (hash >> 15));
    }

    public static RtfWord Lookup(ReadOnlySpan<byte> name)
    {
        int slot = Hash(name) & (TableSize - 1);
        while (Names[slot] is { } existing)
        {
            if (existing.AsSpan().SequenceEqual(name)) return Words[slot];
            slot = (slot + 1) & (TableSize - 1);
        }
        return RtfWord.Unknown;
    }
}

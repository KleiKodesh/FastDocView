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
    FootnoteStart, EndnoteStart, RightToLeftDocument, AutoHyphenation, DefaultLanguage,
    //  page setup (document defaults, then per section) 
    PaperWidth, LeftMargin, RightMargin, Gutter, SectionDefault, SectionPageWidth, SectionLeftMargin, SectionRightMargin, SectionGutter, RightToLeftGutter, GutterAtTop,
    // ── destinations rendered ──
    Footnote, FootnoteIsEndnote, Field, FieldResult, FieldInstruction, FieldLock, FieldPrivate, FormField, FormFieldType, FormFieldResult,
    FormFieldDefaultResult, FormFieldListEntry, FormFieldDefaultText, BookmarkStart, BookmarkEnd, ListText, ParagraphNumberText, UnicodePair, UnicodeDestination,
    Picture, ShapePicture, NonShapePicture, Shape, ShapeGroup, ShapeInstruction, ShapeResult, ShapeText, ShapeProperty, ShapePropertyName, ShapePropertyValue,
    ShapeLeft, ShapeTop, ShapeRight, ShapeBottom, ShapeWrap, ShapeWrapSide, ObjectGroup, MathGroup, NestedTableProperties,
    // ── unicode / special characters ──
    UnicodeSkipCount, UnicodeCharacter, Binary, IgnorableDestination, HexCharacter,
    Paragraph, Section, Page, Line, Tab, Cell, Row, NestedCell, NestedRow, Column,
    EmDash, EnDash, Bullet, LeftQuote, RightQuote, LeftDoubleQuote, RightDoubleQuote, EmSpace, EnSpace, QuarterEmSpace,
    ZeroWidthJoiner, ZeroWidthNonJoiner, ZeroWidthBreakOpportunity, ZeroWidthNoBreak, LeftToRightMark, RightToLeftMark,
    NonBreakingSpace, OptionalHyphen, NonBreakingHyphen, Backslash, OpenBrace, CloseBrace, FootnoteMark, SubEntry, Formula,
    // ── paragraph ──
    ParagraphDefault, AlignLeft, AlignRight, AlignCenter, AlignJustify, AlignDistribute,
    LeftIndent, RightIndent, StartIndent, EndIndent, FirstLineIndent, SpaceBefore, SpaceAfter, LineSpacing, LineSpacingMultiple,
    RightToLeftParagraph, LeftToRightParagraph, InTable, NestingLevel, PageBreakBefore, DropCapLines, DropCapType, FrameHorizontalDistance, ParagraphHyphenation,
    // ── table row ──
    RowDefault, CellBoundary, CellShading, RightToLeftRow, LeftToRightRow, RowLeft, RowAlignLeft, RowAlignCenter, RowAlignRight,
    // ── character ──
    Plain, Bold, Italic, AssociatedBold, AssociatedItalic, Underline, UnderlineDouble, UnderlineDotted, UnderlineDashed, UnderlineWave,
    UnderlineThick, UnderlineWords, UnderlineNone, Strike, StrikeDouble, FontSize, AssociatedFontSize, ForegroundColor, BackgroundColor,
    CharacterShading, Highlight, Superscript, Subscript, NoSuperSubscript, RaisedPosition, LoweredPosition, Hidden, AllCaps, SmallCaps,
    Outline, Shadow, Emboss, Imprint, CharacterSpacing, CharacterScale, RightToLeftCharacters, LeftToRightCharacters, LineBreakClear,
    UnderlineColor, CharacterPatternColor, CharacterShadingPercent, CharacterShadingPattern, Language,
    EmphasisNone, EmphasisDot, EmphasisComma, EmphasisCircle, EmphasisUnderDot,
    // ── borders and paragraph shading ──
    CharacterBorder, ParagraphBorderTop, ParagraphBorderLeft, ParagraphBorderBottom, ParagraphBorderRight, ParagraphBorderBox, OtherBorder,
    BorderWidth, BorderColor, BorderSingle, BorderDouble, BorderDotted, BorderDashed, BorderNone,
    ParagraphBackgroundColor, ParagraphPatternColor, ParagraphShadingPercent, ParagraphShadingPattern,
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
        ("ftnstart", RtfWord.FootnoteStart), ("aftnstart", RtfWord.EndnoteStart), ("rtldoc", RtfWord.RightToLeftDocument), ("hyphauto", RtfWord.AutoHyphenation), ("deflang", RtfWord.DefaultLanguage),
        ("paperw", RtfWord.PaperWidth), ("margl", RtfWord.LeftMargin), ("margr", RtfWord.RightMargin), ("gutter", RtfWord.Gutter),
        ("sectd", RtfWord.SectionDefault), ("pgwsxn", RtfWord.SectionPageWidth), ("marglsxn", RtfWord.SectionLeftMargin),
        ("margrsxn", RtfWord.SectionRightMargin), ("guttersxn", RtfWord.SectionGutter), ("rtlgutter", RtfWord.RightToLeftGutter), ("gutterprl", RtfWord.GutterAtTop),
        ("footnote", RtfWord.Footnote), ("ftnalt", RtfWord.FootnoteIsEndnote), ("field", RtfWord.Field), ("fldrslt", RtfWord.FieldResult),
        ("fldinst", RtfWord.FieldInstruction), ("bkmkstart", RtfWord.BookmarkStart), ("bkmkend", RtfWord.BookmarkEnd),
        // fields: RTF 1.5 spec "Fields" <fieldmod> and "Form Fields" (LibreOffice rtfdispatchvalue.cxx:1426-1455 FFTYPE / FFDEFRES / FFRES)
        ("fldlock", RtfWord.FieldLock), ("fldpriv", RtfWord.FieldPrivate), ("formfield", RtfWord.FormField), ("fftype", RtfWord.FormFieldType),
        ("ffres", RtfWord.FormFieldResult), ("ffdefres", RtfWord.FormFieldDefaultResult), ("ffl", RtfWord.FormFieldListEntry), ("ffdeftext", RtfWord.FormFieldDefaultText),
        ("listtext", RtfWord.ListText), ("pntext", RtfWord.ParagraphNumberText), ("upr", RtfWord.UnicodePair), ("ud", RtfWord.UnicodeDestination),
        ("pict", RtfWord.Picture), ("shppict", RtfWord.ShapePicture), ("nonshppict", RtfWord.NonShapePicture), ("shp", RtfWord.Shape),
        ("shpgrp", RtfWord.ShapeGroup), ("shpinst", RtfWord.ShapeInstruction), ("shprslt", RtfWord.ShapeResult), ("shptxt", RtfWord.ShapeText),
        ("sp", RtfWord.ShapeProperty), ("sn", RtfWord.ShapePropertyName), ("sv", RtfWord.ShapePropertyValue), ("shpleft", RtfWord.ShapeLeft),
        ("shptop", RtfWord.ShapeTop), ("shpright", RtfWord.ShapeRight), ("shpbottom", RtfWord.ShapeBottom), ("shpwr", RtfWord.ShapeWrap), ("shpwrk", RtfWord.ShapeWrapSide),
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
        ("pagebb", RtfWord.PageBreakBefore), ("dropcapli", RtfWord.DropCapLines), ("dropcapt", RtfWord.DropCapType), ("dxfrtext", RtfWord.FrameHorizontalDistance), ("dfrmtxtx", RtfWord.FrameHorizontalDistance), ("hyphpar", RtfWord.ParagraphHyphenation), ("lang", RtfWord.Language),
        ("trowd", RtfWord.RowDefault), ("cellx", RtfWord.CellBoundary), ("clcbpat", RtfWord.CellShading), ("rtlrow", RtfWord.RightToLeftRow), ("ltrrow", RtfWord.LeftToRightRow),
        ("trleft", RtfWord.RowLeft), ("trql", RtfWord.RowAlignLeft), ("trqc", RtfWord.RowAlignCenter), ("trqr", RtfWord.RowAlignRight),
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
        ("impr", RtfWord.Imprint), ("expndtw", RtfWord.CharacterSpacing), ("charscalex", RtfWord.CharacterScale), ("rtlch", RtfWord.RightToLeftCharacters), ("ltrch", RtfWord.LeftToRightCharacters),
        ("lbr", RtfWord.LineBreakClear),
        ("ulc", RtfWord.UnderlineColor), ("chcfpat", RtfWord.CharacterPatternColor), ("chshdng", RtfWord.CharacterShadingPercent),
        ("chbghoriz", RtfWord.CharacterShadingPattern), ("chbgvert", RtfWord.CharacterShadingPattern), ("chbgfdiag", RtfWord.CharacterShadingPattern),
        ("chbgbdiag", RtfWord.CharacterShadingPattern), ("chbgcross", RtfWord.CharacterShadingPattern), ("chbgdcross", RtfWord.CharacterShadingPattern),
        ("chbgdkhoriz", RtfWord.CharacterShadingPattern), ("chbgdkvert", RtfWord.CharacterShadingPattern), ("chbgdkfdiag", RtfWord.CharacterShadingPattern),
        ("chbgdkbdiag", RtfWord.CharacterShadingPattern), ("chbgdkcross", RtfWord.CharacterShadingPattern), ("chbgdkdcross", RtfWord.CharacterShadingPattern),
        ("accnone", RtfWord.EmphasisNone), ("accdot", RtfWord.EmphasisDot), ("acccomma", RtfWord.EmphasisComma), ("acccircle", RtfWord.EmphasisCircle),
        ("accunderdot", RtfWord.EmphasisUnderDot),
        ("chbrdr", RtfWord.CharacterBorder), ("brdrt", RtfWord.ParagraphBorderTop), ("brdrl", RtfWord.ParagraphBorderLeft), ("brdrb", RtfWord.ParagraphBorderBottom),
        ("brdrr", RtfWord.ParagraphBorderRight), ("box", RtfWord.ParagraphBorderBox),
        // borders this reader does not draw: the \brdr* modifiers after them must not reach a paragraph or character border
        ("brdrbtw", RtfWord.OtherBorder), ("brdrbar", RtfWord.OtherBorder), ("clbrdrt", RtfWord.OtherBorder), ("clbrdrl", RtfWord.OtherBorder),
        ("clbrdrb", RtfWord.OtherBorder), ("clbrdrr", RtfWord.OtherBorder), ("pgbrdrt", RtfWord.OtherBorder), ("pgbrdrl", RtfWord.OtherBorder),
        ("pgbrdrb", RtfWord.OtherBorder), ("pgbrdrr", RtfWord.OtherBorder), ("trbrdrt", RtfWord.OtherBorder), ("trbrdrl", RtfWord.OtherBorder),
        ("trbrdrb", RtfWord.OtherBorder), ("trbrdrr", RtfWord.OtherBorder), ("trbrdrh", RtfWord.OtherBorder), ("trbrdrv", RtfWord.OtherBorder),
        ("brdrw", RtfWord.BorderWidth), ("brdrcf", RtfWord.BorderColor),
        ("brdrs", RtfWord.BorderSingle), ("brdrhair", RtfWord.BorderSingle), ("brdrth", RtfWord.BorderSingle), ("brdrtriple", RtfWord.BorderSingle),
        ("brdrtnthsg", RtfWord.BorderSingle), ("brdrthtnsg", RtfWord.BorderSingle), ("brdrtnthtnsg", RtfWord.BorderSingle), ("brdrtnthmg", RtfWord.BorderSingle),
        ("brdrthtnmg", RtfWord.BorderSingle), ("brdrtnthtnmg", RtfWord.BorderSingle), ("brdrtnthlg", RtfWord.BorderSingle), ("brdrthtnlg", RtfWord.BorderSingle),
        ("brdrtnthtnlg", RtfWord.BorderSingle), ("brdrwavy", RtfWord.BorderSingle), ("brdrwavydb", RtfWord.BorderSingle), ("brdrdashdotstr", RtfWord.BorderSingle),
        ("brdremboss", RtfWord.BorderSingle), ("brdrengrave", RtfWord.BorderSingle), ("brdroutset", RtfWord.BorderSingle), ("brdrinset", RtfWord.BorderSingle),
        ("brdrdashd", RtfWord.BorderSingle), ("brdrdashdd", RtfWord.BorderSingle), ("brdrframe", RtfWord.BorderSingle),
        ("brdrdb", RtfWord.BorderDouble), ("brdrdot", RtfWord.BorderDotted), ("brdrdash", RtfWord.BorderDashed), ("brdrdashsm", RtfWord.BorderDashed),
        ("brdrnone", RtfWord.BorderNone), ("brdrnil", RtfWord.BorderNone),
        ("cbpat", RtfWord.ParagraphBackgroundColor), ("cfpat", RtfWord.ParagraphPatternColor), ("shading", RtfWord.ParagraphShadingPercent),
        ("bghoriz", RtfWord.ParagraphShadingPattern), ("bgvert", RtfWord.ParagraphShadingPattern), ("bgfdiag", RtfWord.ParagraphShadingPattern),
        ("bgbdiag", RtfWord.ParagraphShadingPattern), ("bgcross", RtfWord.ParagraphShadingPattern), ("bgdcross", RtfWord.ParagraphShadingPattern),
        ("bgdkhoriz", RtfWord.ParagraphShadingPattern), ("bgdkvert", RtfWord.ParagraphShadingPattern), ("bgdkfdiag", RtfWord.ParagraphShadingPattern),
        ("bgdkbdiag", RtfWord.ParagraphShadingPattern), ("bgdkcross", RtfWord.ParagraphShadingPattern), ("bgdkdcross", RtfWord.ParagraphShadingPattern),
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

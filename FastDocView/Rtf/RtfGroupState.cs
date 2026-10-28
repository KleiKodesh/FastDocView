using System.Runtime.InteropServices;

namespace FastDocView.Rtf;

/// <summary>Where the text of a group goes.</summary>
internal enum RtfDestination : byte
{
    Body = 0,
    Skip,
    FontTable,
    ColorTable,
    StyleSheet,
    ListTable,
    List,
    ListLevel,
    ListLevelText,
    ListLevelNumbers,
    ListOverrideTable,
    ListOverride,
    ListOverrideLevel,
    Footnote,
    ListText,
}

/// <summary>
/// Which character slot (Latin or complex script) \f, \fs, \b, \i and their associated forms set — LibreOffice's
/// RTFParserState::RunType (rtfdispatchflag.cxx LTRCH / RTLCH, rtfdispatchvalue.cxx F / AF / FS / AFS):
/// after \rtlch (or \ltrch then \rtlch) they set the complex-script slot, otherwise the Latin one.
/// </summary>
internal enum RtfRunType : byte { None, RightToLeftFirst, RightToLeftThenLeftToRight, LeftToRightFirst, LeftToRightThenRightToLeft }

[Flags]
internal enum RtfStateFlags : uint
{
    None = 0,
    Bold = 1 << 0, Italic = 1 << 1, ComplexBold = 1 << 2, ComplexItalic = 1 << 3, Strike = 1 << 4, Hidden = 1 << 5,
    AllCaps = 1 << 6, SmallCaps = 1 << 7, Superscript = 1 << 8, Subscript = 1 << 9, Outline = 1 << 10, Shadow = 1 << 11,
    Emboss = 1 << 12, Imprint = 1 << 13,
    // the complex-script slot was set explicitly since \plain (RichEdit writes Hebrew with \f / \fs only, no \rtlch)
    ComplexFontSet = 1 << 14, ComplexSizeSet = 1 << 15, ComplexBoldSet = 1 << 16, ComplexItalicSet = 1 << 17,
    CharacterMask = (1 << 20) - 1,
    // paragraph
    RightToLeftParagraph = 1 << 20, InTable = 1 << 21, PageBreakBefore = 1 << 22, LineSpacingMultiple = 1 << 23,
    HasStartIndent = 1 << 24, HasEndIndent = 1 << 25,
    ParagraphMask = 0xFFu << 20,
    // group
    DropText = 1u << 30,          // the ANSI half of \upr (the \ud half is used)
}

internal enum RtfUnderline : byte { None, Single, Double, Dotted, Dashed, Wave, Thick, Words }

internal enum RtfAlignment : byte { Unset, Left, Right, Center, Justify }

/// <summary>
/// The state of one RTF group (RTF 1.5 spec "Conventions of an RTF reader": destination, character, paragraph
/// properties live in the group state; '{' saves it, '}' restores it). Plain unmanaged data so the index can
/// store a stack snapshot per block as raw bytes and a window restores it without re-reading the file from the start.
/// </summary>
[StructLayout(LayoutKind.Sequential, Pack = 1)]
internal struct RtfGroupState
{
    public RtfDestination Destination;
    public RtfRunType RunType;
    public byte UnicodeSkipCount;          // \ucN (default 1)
    public RtfUnderline Underline;
    public RtfAlignment Alignment;
    public byte ListLevel;                 // \ilvl
    public sbyte OutlineLevel;             // \outlinelevel, -1 = none
    public byte NestingLevel;              // \itap
    public RtfStateFlags Flags;
    public ushort CodePage;                // the code page \'xx bytes decode with
    public short Font, ComplexFont;        // \f numbers (-1 none)
    public short FontSize, ComplexFontSize;   // half-points
    public short ForegroundColor, BackgroundColor, ShadingColor;   // color table indexes, 0 = auto
    public byte Highlight;                 // spec palette 1..16
    public short Position;                 // \up (+) / \dn (-) half-points
    public short CharacterSpacing;         // \expndtw twips
    public short Style;                    // \s
    public short ListOverride;             // \ls (0 = none)
    public int LeftIndent, RightIndent, StartIndent, EndIndent, FirstLineIndent, SpaceBefore, SpaceAfter, LineSpacing;

    public static int Size => System.Runtime.CompilerServices.Unsafe.SizeOf<RtfGroupState>();

    public readonly bool Has(RtfStateFlags flag) => (Flags & flag) != 0;

    public void Set(RtfStateFlags flag, bool on)
    {
        if (on) Flags |= flag; else Flags &= ~flag;
    }

    public bool IsInTable => Has(RtfStateFlags.InTable) || NestingLevel > 0;

    /// <summary>Character properties back to the document defaults (\plain).</summary>
    public void ResetCharacter(short defaultFont, ushort defaultCodePage)
    {
        Flags &= ~RtfStateFlags.CharacterMask;
        RunType = RtfRunType.None;
        Underline = RtfUnderline.None;
        Font = defaultFont; ComplexFont = defaultFont;
        FontSize = 24; ComplexFontSize = 24;     // \fs default 24 (RTF 1.5 spec, LibreOffice table)
        ForegroundColor = 0; BackgroundColor = 0; ShadingColor = 0; Highlight = 0;
        Position = 0; CharacterSpacing = 0;
        CodePage = defaultCodePage;
    }

    /// <summary>Paragraph properties back to the defaults (\pard; LibreOffice also selects style 0).</summary>
    public void ResetParagraph()
    {
        Flags &= ~RtfStateFlags.ParagraphMask;
        Alignment = RtfAlignment.Unset;
        ListLevel = 0; ListOverride = 0; OutlineLevel = -1; NestingLevel = 0; Style = 0;
        LeftIndent = RightIndent = StartIndent = EndIndent = FirstLineIndent = SpaceBefore = SpaceAfter = LineSpacing = 0;
    }

    public static RtfGroupState Initial(short defaultFont, ushort defaultCodePage)
    {
        var state = new RtfGroupState { UnicodeSkipCount = 1 };
        state.ResetCharacter(defaultFont, defaultCodePage);
        state.ResetParagraph();
        return state;
    }
}

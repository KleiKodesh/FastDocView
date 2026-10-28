namespace FastDocView.Html;

/// <summary>The toggle (on/off) run properties of ECMA-376 Part 1 §17.7.3: b, bCs, caps, emboss, i, iCs, imprint, outline, shadow, smallCaps, strike, vanish.</summary>
[Flags]
public enum WordToggleProperty : ushort
{
    None = 0,
    Bold = 1 << 0,
    BoldComplexScript = 1 << 1,
    Italic = 1 << 2,
    ItalicComplexScript = 1 << 3,
    Caps = 1 << 4,
    SmallCaps = 1 << 5,
    Strike = 1 << 6,
    Emboss = 1 << 7,
    Imprint = 1 << 8,
    Outline = 1 << 9,
    Shadow = 1 << 10,
    Vanish = 1 << 11,
}

/// <summary>Where a run property comes from in the .docx style hierarchy (§17.7.2), lowest priority first.</summary>
public enum WordStyleLayer { DocumentDefaults, TableStyle, ParagraphStyle, CharacterStyle, Direct }

/// <summary>
/// The toggle properties of a run, kept per style layer so they can be combined the way Word does.
/// Checked in Word 16 (synthetic documents read back through COM and exported to PDF, 2026-10-28):
///   • inside one style's w:basedOn chain the last value wins (a paragraph style is absolute — [MS-OI29500]
///     2.1.258 on §17.7.8 — and Word treats a character style chain the same way);
///   • the table, paragraph and character style results are XORed (bold paragraph style + bold character
///     style = not bold; a value of false contributes nothing);
///   • a toggle on in the document defaults stays on whatever the styles say;
///   • direct formatting is absolute (§17.7.3).
/// </summary>
public struct WordToggleLayers
{
    private ushort _defaults, _table, _paragraph, _character, _directSet, _directOn;

    public void Set(WordStyleLayer layer, WordToggleProperty property, bool isOn)
    {
        ushort bit = (ushort)property;
        switch (layer)
        {
            case WordStyleLayer.DocumentDefaults: _defaults = With(_defaults, bit, isOn); break;
            case WordStyleLayer.TableStyle: _table = With(_table, bit, isOn); break;
            case WordStyleLayer.ParagraphStyle: _paragraph = With(_paragraph, bit, isOn); break;
            case WordStyleLayer.CharacterStyle: _character = With(_character, bit, isOn); break;
            default: _directSet |= bit; _directOn = With(_directOn, bit, isOn); break;
        }
    }

    public readonly bool IsOn(WordToggleProperty property)
    {
        ushort bit = (ushort)property;
        if ((_directSet & bit) != 0) return (_directOn & bit) != 0;
        return ((_defaults | (_table ^ _paragraph ^ _character)) & bit) != 0;
    }

    private static ushort With(ushort bits, ushort bit, bool isOn) => (ushort)(isOn ? bits | bit : bits & ~bit);
}

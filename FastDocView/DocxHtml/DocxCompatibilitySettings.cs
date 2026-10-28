using System.Globalization;
using System.Xml.Linq;

namespace FastDocView.DocxHtml;

/// <summary>
/// settings.xml w:compat (ECMA-376 §17.15.3): the compatibility options the renderer acts on, read the way LibreOffice's
/// SettingsTable.cxx reads them. compatibilityMode comes from w:compatSetting name="compatibilityMode"
/// uri="http://schemas.microsoft.com/office/word"; when repeated the highest wins and only values above 10 count
/// (GetWordCompatibilityMode); absent means Word 2007 (12). The on/off options are CT_OnOff (present = on unless val is
/// false). settings.xml is a small part already parsed at every open (theme mapping, note formats), so these live with
/// the document parts and are not copied into the body index.
/// </summary>
public sealed class DocxCompatibilitySettings
{
    private static readonly XNamespace W = DocxStyleCascade.W;
    private const string WordCompatibilityUri = "http://schemas.microsoft.com/office/word";

    /// <summary>The Word version the document is laid out for: 11 (2003), 12 (2007, also when absent), 14 (2010), 15 (2013+).</summary>
    public int CompatibilityMode { get; } = 12;

    /// <summary>w:doNotExpandShiftReturn: a justified line ending in a manual line break is NOT stretched (Word's default stretches it).</summary>
    public bool DoesNotExpandShiftReturn { get; }

    /// <summary>w:doNotUseHTMLParagraphAutoSpacing: auto spacing before / after is 100 twips instead of 280 (DomainMapper.cxx).</summary>
    public bool DoesNotUseHtmlParagraphAutoSpacing { get; }

    /// <summary>w:ulTrailSpace: underline the spaces at the end of a line too (Word's default leaves them plain).</summary>
    public bool UnderlinesTrailingSpaces { get; }

    /// <summary>compatSetting overrideTableStyleFontSizeAndJustification: the default paragraph style's 11 / 12 pt size and left alignment are not replaced by a table style's.</summary>
    public bool OverridesTableStyleFontSizeAndJustification { get; }

    /// <summary>Every w:compat child and compatSetting name present (for the formatting audit).</summary>
    public IReadOnlyList<string> PresentOptions { get; } = [];

    public static readonly DocxCompatibilitySettings Default = new(null);

    public DocxCompatibilitySettings(XDocument? settingsDocument)
    {
        XElement? compat = settingsDocument?.Root?.Element(W + "compat");
        if (compat is null) return;
        var present = new List<string>();
        int mode = -1;
        foreach (XElement option in compat.Elements())
        {
            if (option.Name == W + "compatSetting")
            {
                string? name = (string?)option.Attribute(W + "name");
                string? value = (string?)option.Attribute(W + "val");
                if (name is null) continue;
                present.Add(name);
                if ((string?)option.Attribute(W + "uri") != WordCompatibilityUri) continue;
                if (name == "compatibilityMode" && int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int candidate) && candidate > 10 && candidate > mode)
                    mode = candidate;
                else if (name == "overrideTableStyleFontSizeAndJustification") OverridesTableStyleFontSizeAndJustification = value is "1" or "true" or "on";
                continue;
            }
            if (option.Name.Namespace != W) continue;
            present.Add(option.Name.LocalName);
            bool isOn = DocxStyleCascade.IsOn(option.Attribute(W + "val"));
            switch (option.Name.LocalName)
            {
                case "doNotExpandShiftReturn": DoesNotExpandShiftReturn = isOn; break;
                case "doNotUseHTMLParagraphAutoSpacing": DoesNotUseHtmlParagraphAutoSpacing = isOn; break;
                case "ulTrailSpace": UnderlinesTrailingSpaces = isOn; break;
            }
        }
        if (mode > 0) CompatibilityMode = mode;
        PresentOptions = present;
    }
}

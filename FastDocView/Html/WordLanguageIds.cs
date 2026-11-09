namespace FastDocView.Html;

/// <summary>
/// Windows language ids (LCID, [MS-LCID]) as BCP 47 tags, for the languages whose text Chromium can hyphenate plus the
/// right-to-left ones Word documents here use. .doc (sprmCRgLid0 / sprmCRgLid0_80) and RTF (\lang) store an LCID; the tag
/// only goes on a hyphenated paragraph's lang attribute. A table rather than CultureInfo: the library also runs with
/// invariant globalization (AOT / WASM), where LCID lookups fail.
/// </summary>
public static class WordLanguageIds
{
    private static readonly Dictionary<int, string> Tags = new()
    {
        [0x0409] = "en-US", [0x0809] = "en-GB", [0x0C09] = "en-AU", [0x1009] = "en-CA", [0x1409] = "en-NZ", [0x1809] = "en-IE", [0x1C09] = "en-ZA",
        [0x0407] = "de-DE", [0x0807] = "de-CH", [0x0C07] = "de-AT", [0x040C] = "fr-FR", [0x080C] = "fr-BE", [0x0C0C] = "fr-CA", [0x100C] = "fr-CH",
        [0x040A] = "es-ES", [0x0C0A] = "es-ES", [0x080A] = "es-MX", [0x0410] = "it-IT", [0x0413] = "nl-NL", [0x0813] = "nl-BE",
        [0x0416] = "pt-BR", [0x0816] = "pt-PT", [0x0419] = "ru-RU", [0x0422] = "uk-UA", [0x0415] = "pl-PL", [0x0405] = "cs-CZ", [0x041B] = "sk-SK",
        [0x041D] = "sv-SE", [0x0406] = "da-DK", [0x0414] = "nb-NO", [0x0814] = "nn-NO", [0x040B] = "fi-FI", [0x040E] = "hu-HU", [0x0418] = "ro-RO",
        [0x0408] = "el-GR", [0x041F] = "tr-TR", [0x0402] = "bg-BG", [0x041A] = "hr-HR", [0x0424] = "sl-SI", [0x0425] = "et-EE", [0x0426] = "lv-LV",
        [0x0427] = "lt-LT", [0x0403] = "ca-ES", [0x0476] = "la", [0x040D] = "he-IL", [0x0401] = "ar-SA", [0x043D] = "yi",
    };

    /// <summary>The BCP 47 tag of a language id, or null when it is not in the table (0x0400 "none", 0x0000, unknown ids).</summary>
    public static string? TagOf(int languageId) => Tags.TryGetValue(languageId, out string? tag) ? tag : null;
}

namespace FastDocView.Html;

/// <summary>
/// Links into a Jewish-library app: Otzaria (otzaria://, and its book://&lt;title&gt;#heading links), Zayit (zayit://, and
/// its seforim://line/&lt;id&gt; citation links) and KitveiHakodesh's own (kitveihakodeshapp://, formerly seforimapp://).
/// The host opens these inside itself, never in a browser, so every reader marks them (data-app-link) the same way.
/// The one place the schemes are listed - by scheme only, never by parsing the rest: a book title in the host part
/// (book://&lt;Hebrew title with spaces&gt;) is not a valid URI host, but is still a link the host can open.
/// </summary>
internal static class WordLibraryAppLink
{
    private static readonly System.Text.RegularExpressions.Regex ZayitWebOpen = new(@"^https?://(?:www\.)?zayitapp\.com/(?:[a-z]{2}/)?open(?:\.html)?/?#.", System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.CultureInvariant);

    private static readonly string[] Schemes = ["otzaria", "zayit", "kitveihakodeshapp", "seforimapp", "book", "seforim"];

    /// <summary>True when the target (already HTML-decoded) starts with one of the library apps' schemes followed by ':'.</summary>
    internal static bool IsLibraryAppLink(string? target)
    {
        if (target is null) return false;
        string href = target.Trim();
        // Zayit's web share page hands its fragment to zayit:// (Zayit website/public/open.html)
        if (ZayitWebOpen.IsMatch(href)) return true;
        int colon = href.IndexOf(':');
        if (colon <= 0) return false;
        foreach (string scheme in Schemes)
            if (colon == scheme.Length && string.Compare(href, 0, scheme, 0, colon, System.StringComparison.OrdinalIgnoreCase) == 0) return true;
        return false;
    }
}

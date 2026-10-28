using System.Xml.Linq;

namespace WordDocumentRandomAccess.DocxHtml;

/// <summary>
/// word/numbering.xml (ECMA-376 §17.9): numId → abstractNumId → per-level format and bullet text.
/// Windowed rendering cannot count list items that were never rendered, so numbered levels get a
/// neutral marker instead of a running number (see NOTES.md, "Numbering").
/// </summary>
public sealed class DocxNumberingDefinitions
{
    private static readonly XNamespace W = DocxStyleCascade.W;

    private readonly Dictionary<int, int> _abstractIdByNumId = new();
    private readonly Dictionary<(int abstractId, int level), (string format, string text)> _levels = new();

    public DocxNumberingDefinitions(XDocument? numberingDocument)
    {
        XElement? root = numberingDocument?.Root;
        if (root is null) return;
        foreach (XElement num in root.Elements(W + "num"))
        {
            int? numId = DocxStyleCascade.TwipsAttribute(num, "numId");
            int? abstractId = DocxStyleCascade.TwipsAttribute(num.Element(W + "abstractNumId"), "val");
            if (numId is { } id && abstractId is { } abstractNumId) _abstractIdByNumId[id] = abstractNumId;
        }
        foreach (XElement abstractNum in root.Elements(W + "abstractNum"))
        {
            int? abstractId = DocxStyleCascade.TwipsAttribute(abstractNum, "abstractNumId");
            if (abstractId is not { } id) continue;
            foreach (XElement level in abstractNum.Elements(W + "lvl"))
            {
                int levelIndex = DocxStyleCascade.TwipsAttribute(level, "ilvl") ?? 0;
                string format = (string?)level.Element(W + "numFmt")?.Attribute(W + "val") ?? "decimal";
                string text = (string?)level.Element(W + "lvlText")?.Attribute(W + "val") ?? "";
                _levels[(id, levelIndex)] = (format, text);
            }
        }
    }

    /// <summary>The marker shown before a list paragraph, or null when the paragraph is not in a list.</summary>
    public string? MarkerFor(int numId, int level)
    {
        if (numId <= 0 || !_abstractIdByNumId.TryGetValue(numId, out int abstractId)) return null;
        if (!_levels.TryGetValue((abstractId, level), out var definition)) return "•";
        return definition.format switch
        {
            "bullet" => definition.text.Length == 1 && definition.text[0] is >= '' and <= '' ? "•" : definition.text,
            "none" => "",
            _ => "–",
        };
    }
}

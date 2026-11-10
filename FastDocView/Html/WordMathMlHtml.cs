using System.Text;
using System.Xml.Linq;

namespace FastDocView.Html;

/// <summary>
/// A MathML formula (an OpenDocument formula object: LibreOffice Math stores MathML 2 / 3 in the object's content.xml,
/// starmath/source/mathml/mathmlexport.cxx; ODF 1.3 Part 3 section 16 "math:math") written out as MathML Core for
/// Chromium. Only MathML elements and presentation attributes are copied (anything else is dropped, so a document cannot
/// inject markup); semantics shows its first child (the annotation holds the StarMath source); mstyle / mathvariant / mathcolor
/// that MathML Core does not have are mapped: mathvariant other than "normal" to the Mathematical Alphanumeric Symbols
/// (<see cref="WordOfficeMathHtml.StyledText"/>), mathcolor / mathbackground / mathsize to CSS, mfenced to an mrow of fences,
/// menclose to a bordered mrow. LibreOffice draws its formula objects in display style (starmath/source/document.cxx), so
/// the formula is an inline math element in display style.
/// </summary>
public static class WordMathMlHtml
{
    public const string MathMlNamespace = "http://www.w3.org/1998/Math/MathML";

    private static readonly HashSet<string> Elements = new(StringComparer.Ordinal)
    {
        "mrow", "mi", "mn", "mo", "mtext", "ms", "mspace", "mfrac", "msqrt", "mroot", "msub", "msup", "msubsup", "munder", "mover",
        "munderover", "mmultiscripts", "mprescripts", "none", "mtable", "mtr", "mtd", "mphantom", "mpadded", "merror", "mstyle",
        "mlabeledtr", "maligngroup", "malignmark",
    };

    private static readonly HashSet<string> Attributes = new(StringComparer.Ordinal)
    {
        "displaystyle", "scriptlevel", "linethickness", "stretchy", "symmetric", "largeop", "movablelimits", "accent", "accentunder",
        "fence", "separator", "form", "lspace", "rspace", "minsize", "maxsize", "width", "height", "depth", "voffset",
    };

    /// <summary>Appends a math:math element (null or not MathML: nothing is written, false).</summary>
    public static bool Append(StringBuilder html, XElement? math)
    {
        if (math is null || math.Name.NamespaceName != MathMlNamespace || math.Name.LocalName != "math") return false;
        html.Append("<math class=\"odf-formula\" dir=\"ltr\" displaystyle=\"true\" style=\"math-style:normal\">");
        foreach (XElement child in math.Elements()) AppendElement(html, child, null, null);
        html.Append("</math>");
        return true;
    }

    /// <summary>The text a formula shows (its tokens' text), for search.</summary>
    public static string Text(XElement math)
    {
        var text = new StringBuilder();
        foreach (XElement token in math.Descendants())
            if (token.Name.NamespaceName == MathMlNamespace && token.Name.LocalName is "mi" or "mn" or "mo" or "mtext" or "ms" && !token.Ancestors().Any(a => a.Name.LocalName == "annotation" || a.Name.LocalName == "annotation-xml"))
                text.Append(token.Value);
        return text.ToString();
    }

    private static void AppendElement(StringBuilder html, XElement element, string? variant, string? css)
    {
        if (element.Name.NamespaceName != MathMlNamespace) return;
        string name = element.Name.LocalName;
        switch (name)
        {
            case "semantics":
                if (element.Elements().FirstOrDefault() is { } shown && shown.Name.LocalName is not ("annotation" or "annotation-xml")) AppendElement(html, shown, variant, css);
                return;
            case "annotation": case "annotation-xml": return;
            case "mfenced":
            {
                // MathML 2 mfenced (not in MathML Core): open, the children split by separators, close
                string open = (string?)element.Attribute("open") ?? "(", close = (string?)element.Attribute("close") ?? ")";
                string separators = ((string?)element.Attribute("separators") ?? ",").Replace(" ", "");
                html.Append("<mrow>");
                if (open.Length > 0) html.Append("<mo fence=\"true\" form=\"prefix\">").Append(WordHtmlText.Encode(open)).Append("</mo>");
                int index = 0;
                foreach (XElement child in element.Elements())
                {
                    if (index > 0 && separators.Length > 0) html.Append("<mo separator=\"true\">").Append(WordHtmlText.Encode(separators[Math.Min(index - 1, separators.Length - 1)].ToString())).Append("</mo>");
                    AppendElement(html, child, variant, css);
                    index++;
                }
                if (close.Length > 0) html.Append("<mo fence=\"true\" form=\"postfix\">").Append(WordHtmlText.Encode(close)).Append("</mo>");
                html.Append("</mrow>");
                return;
            }
            case "menclose":
                html.Append("<mrow style=\"padding:0.1em;border:1px solid currentColor\">");
                foreach (XElement child in element.Elements()) AppendElement(html, child, variant, css);
                html.Append("</mrow>");
                return;
        }
        if (!Elements.Contains(name)) return;
        // mathvariant on mstyle reaches its tokens; on a token it is its own
        string? ownVariant = (string?)element.Attribute("mathvariant") ?? variant;
        string? ownCss = Css(element, css);
        bool isToken = name is "mi" or "mn" or "mo" or "mtext" or "ms";
        string tag = name == "mstyle" ? "mrow" : name;
        html.Append('<').Append(tag);
        foreach (XAttribute attribute in element.Attributes())
            if (attribute.Name.NamespaceName.Length == 0 && Attributes.Contains(attribute.Name.LocalName))
                html.Append(' ').Append(attribute.Name.LocalName).Append("=\"").Append(WordHtmlText.Encode(attribute.Value)).Append('"');
        string text = element.Value;
        if (isToken)
        {
            if (ownVariant is not null && ownVariant != "normal" && WordOfficeMathHtml.StyledText(text, ownVariant) is { } styled) { text = styled; html.Append(" mathvariant=\"normal\""); }
            else if (ownVariant == "normal" && name == "mi") html.Append(" mathvariant=\"normal\"");
            if (ownCss is { Length: > 0 }) html.Append(" style=\"").Append(ownCss).Append('"');
            html.Append('>').Append(WordHtmlText.Encode(text)).Append("</").Append(tag).Append('>');
            return;
        }
        if (name == "mstyle" && ownCss is { Length: > 0 } && ownCss != css) html.Append(" style=\"").Append(ownCss).Append('"');
        html.Append('>');
        foreach (XElement child in element.Elements()) AppendElement(html, child, ownVariant, name == "mstyle" ? ownCss : css);
        html.Append("</").Append(tag).Append('>');
    }

    /// <summary>mathcolor / mathbackground / mathsize (and their MathML 1 forms color / background / fontsize) as CSS; inherited from mstyle.</summary>
    private static string? Css(XElement element, string? inherited)
    {
        var css = new StringBuilder();
        string? color = (string?)element.Attribute("mathcolor") ?? (string?)element.Attribute("color");
        string? background = (string?)element.Attribute("mathbackground") ?? (string?)element.Attribute("background");
        string? size = (string?)element.Attribute("mathsize") ?? (string?)element.Attribute("fontsize");
        if (IsSafe(color)) css.Append("color:").Append(color).Append(';');
        if (IsSafe(background)) css.Append("background-color:").Append(background).Append(';');
        if (IsSafe(size)) css.Append("font-size:").Append(size switch { "small" => "80%", "big" => "125%", "normal" => "100%", _ => size }).Append(';');
        if (css.Length == 0) return inherited;
        return (inherited ?? "") + css;
    }

    /// <summary>A CSS value from the document: letters, digits, '#', '.', '%' only (a colour name or hex, a length).</summary>
    private static bool IsSafe(string? value) => value is { Length: > 0 and < 40 } && value.All(c => char.IsLetterOrDigit(c) || c is '#' or '.' or '%' or '-');
}

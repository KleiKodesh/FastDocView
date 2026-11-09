namespace FastDocView.Html;

/// <summary>
/// Which property set (Latin or complex script) a list label draws with. Word 16 measured (PDF export of
/// test-data/word-lists/list-bullets-v1.doc / .docx, right-to-left paragraphs, mark sz 11 / szCs 12, level bold + red):
/// the number "1." draws in the complex-script font, size and bold (Arial 12, not bold - the level set only b),
/// while symbol-font bullets (U+F0B7, U+F0A7) and picture bullets follow the Latin size (11 / 24 pt).
/// That is LibreOffice's rule for a number portion (sw/source/core/text/porfld.cxx:222-316 SwFieldPortion::CheckScript):
/// complex-script text is CTL; otherwise, in a right-to-left paragraph (the number takes the paragraph's direction,
/// :258-260), a label without a strong left-to-right character is CTL too (:270-293, #i89825).
/// </summary>
public static class WordListLabelScript
{
    /// <param name="label">The label as stored (before symbol-font recoding: a symbol code U+F0xx is a strong LTR character).</param>
    public static bool UsesComplexScript(string label, bool isParagraphRightToLeft)
    {
        if (label.Length == 0 || label[0] == WordListLabelHtml.PictureLabelMark) return false;
        bool hasStrongLeftToRight = false;
        foreach (char c in label)
        {
            if (IsComplexScript(c)) return true;
            if (IsStrongLeftToRight(c)) hasStrongLeftToRight = true;
        }
        return isParagraphRightToLeft && !hasStrongLeftToRight;
    }

    /// <summary>Hebrew, Arabic, Syriac, Thaana, NKo, Samaritan..., their presentation forms (i18n ScriptType COMPLEX for the RTL scripts).</summary>
    private static bool IsComplexScript(char c)
        => c is (>= '֐' and <= 'ࣿ') or (>= 'יִ' and <= '﷿') or (>= 'ﹰ' and <= '﻿');

    /// <summary>u_charDirection == U_LEFT_TO_RIGHT: letters of the LTR scripts and the private-use area (symbol-font codes).</summary>
    private static bool IsStrongLeftToRight(char c)
        => c is >= '' and <= '' || (char.IsLetter(c) && !IsComplexScript(c));

    /// <summary>The label's formatting with the property set it draws in chosen (a clone when it changes).</summary>
    public static WordRunFormatting WithScript(WordRunFormatting formatting, bool usesComplexScript)
    {
        if (formatting.UsesComplexScriptProperties == usesComplexScript) return formatting;
        WordRunFormatting label = formatting.Clone();
        label.IsRightToLeft = usesComplexScript;
        label.IsComplexScript = false;
        return label;
    }
}

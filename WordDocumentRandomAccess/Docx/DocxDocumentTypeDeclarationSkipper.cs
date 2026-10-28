namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// Finds the end of a &lt;!DOCTYPE …&gt; for the byte scanners, fed one byte at a time after
/// "&lt;!D". XML 1.0 §2.8: doctypedecl ::= '&lt;!DOCTYPE' S Name (S ExternalID)? S? ('[' intSubset ']' S?)? '&gt;' —
/// the declaration ends at the first '&gt;' outside quotes and outside the internal subset; inside the
/// subset, quoted literals (an ENTITY value or SYSTEM id may hold '&gt;' or ']') and comments are
/// skipped. Nothing is resolved: the declarations are only stepped over (apache-poi
/// ExternalEntityInText.docx declares an external entity this way).
/// </summary>
internal sealed class DocxDocumentTypeDeclarationSkipper
{
    private bool _isInsideSubset;
    private byte _quote;            // 0 = not inside a literal
    private int _commentOpenMatched;  // progress through "<!--" inside the subset
    private bool _isInsideComment;
    private int _commentCloseMatched; // progress through "-->"

    public void Reset() { _isInsideSubset = false; _quote = 0; _commentOpenMatched = 0; _isInsideComment = false; _commentCloseMatched = 0; }

    /// <summary>True when <paramref name="b"/> is the declaration's closing '&gt;'.</summary>
    public bool Feed(byte b)
    {
        if (_isInsideComment)
        {
            if (b == (byte)'-') _commentCloseMatched = Math.Min(2, _commentCloseMatched + 1);
            else if (b == (byte)'>' && _commentCloseMatched == 2) { _isInsideComment = false; _commentCloseMatched = 0; }
            else _commentCloseMatched = 0;
            return false;
        }
        if (_quote != 0)
        {
            if (b == _quote) _quote = 0;
            return false;
        }
        if (_isInsideSubset)
        {
            _commentOpenMatched = (_commentOpenMatched, b) switch
            {
                (_, (byte)'<') => 1,
                (1, (byte)'!') => 2,
                (2, (byte)'-') => 3,
                (3, (byte)'-') => 4,
                _ => 0,
            };
            if (_commentOpenMatched == 4) { _isInsideComment = true; _commentOpenMatched = 0; return false; }
        }
        if (b is (byte)'"' or (byte)'\'') _quote = b;
        else if (b == (byte)'[') _isInsideSubset = true;
        else if (b == (byte)']') _isInsideSubset = false;
        else if (b == (byte)'>' && !_isInsideSubset) return true;
        return false;
    }
}

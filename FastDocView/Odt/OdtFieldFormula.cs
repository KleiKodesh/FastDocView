using System.Globalization;
using System.Text;

namespace FastDocView.Odt;

/// <summary>
/// A field value of an OpenDocument text: a number, or a string (string variables, ODF 1.3 part 3 7.7.1 office:value-type="string").
/// </summary>
internal readonly record struct OdtFieldValue(double Number, string? Text)
{
    public bool IsText => Text is not null;
    public static OdtFieldValue Of(double number) => new(number, null);
    public static OdtFieldValue OfText(string text) => new(0, text);

    /// <summary>A string used in arithmetic: its number when it reads as one, else 0 (sw calc.cxx SwSbxValue::GetDouble).</summary>
    public double AsNumber => Text is null ? Number : double.TryParse(Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) ? parsed : 0;

    /// <summary>The value as LibreOffice shows it with the standard number format: up to 10 decimals, no trailing zeros.</summary>
    public string Display => Text ?? FormatNumber(Number);

    public static string FormatNumber(double number)
    {
        if (double.IsNaN(number) || double.IsInfinity(number)) return "";
        double rounded = Math.Round(number, 10);
        if (rounded == 0) rounded = 0;      // no "-0"
        return rounded.ToString("0.##########", CultureInfo.InvariantCulture);
    }
}

/// <summary>
/// Writer's field formula language (the "ooow:" namespace of text:formula / text:condition, ODF 1.3 part 3 7.7, 19.890),
/// evaluated the way LibreOffice's SwCalc does (sw/source/core/bastyp/calc.cxx): numbers, variables (sequences, variables and
/// user fields share one name space), + - * / ^ %, comparisons (== != &lt; &gt; &lt;= &gt;= and EQ NEQ L G LEQ GEQ), AND / OR / XOR /
/// NOT (also &amp;&amp; || !), parentheses and the common functions (ABS SQRT SIN COS TAN ASIN ACOS ATAN ROUND INT FRAC SIGN
/// MIN MAX MEAN SUM POW). Function arguments are separated by '|' or ';' as in Writer. A formula that does not parse gives
/// null: the field keeps the text it was saved with.
/// </summary>
internal sealed class OdtFieldFormula
{
    private readonly string _text;
    private readonly Func<string, OdtFieldValue?> _variable;
    private int _position;
    private int _depth;                       // nesting bound: a malformed formula must not overflow the stack

    private OdtFieldFormula(string text, Func<string, OdtFieldValue?> variable)
    {
        _text = text;
        _variable = variable;
    }

    /// <summary>Strips the formula namespace prefix LibreOffice writes ("ooow:"; ODF 1.3 part 3 19.890 namespaced formula).</summary>
    public static string StripNamespace(string formula)
    {
        int colon = formula.IndexOf(':');
        if (colon > 0 && colon < 8 && formula.Substring(0, colon).All(char.IsLetter) && !(colon + 1 < formula.Length && formula[colon + 1] == '='))
            return formula.Substring(colon + 1);
        return formula;
    }

    public static OdtFieldValue? Evaluate(string formula, Func<string, OdtFieldValue?> variable)
    {
        var parser = new OdtFieldFormula(StripNamespace(formula), variable);
        try
        {
            OdtFieldValue value = parser.Or();
            parser.SkipBlanks();
            return parser._position == parser._text.Length ? value : null;
        }
        catch (FormatException) { return null; }
        catch (OverflowException) { return null; }
    }

    private static readonly FormatException Bad = new("formula");

    private void SkipBlanks() { while (_position < _text.Length && char.IsWhiteSpace(_text[_position])) _position++; }

    private bool Take(string token, bool isWord = false)
    {
        SkipBlanks();
        if (string.Compare(_text, _position, token, 0, token.Length, StringComparison.OrdinalIgnoreCase) != 0) return false;
        int end = _position + token.Length;
        if (isWord && end < _text.Length && (char.IsLetterOrDigit(_text[end]) || _text[end] == '_')) return false;
        _position = end;
        return true;
    }

    private OdtFieldValue Or()
    {
        if (++_depth > 256) throw Bad;
        try { return OrInner(); } finally { _depth--; }
    }

    private OdtFieldValue OrInner()
    {
        OdtFieldValue left = And();
        while (true)
        {
            if (Take("||") || Take("OR", true)) { OdtFieldValue right = And(); left = Bool(Truth(left) || Truth(right)); }
            else if (Take("XOR", true)) { OdtFieldValue right = And(); left = Bool(Truth(left) ^ Truth(right)); }
            else return left;
        }
    }

    private OdtFieldValue And()
    {
        OdtFieldValue left = Comparison();
        while (Take("&&") || Take("AND", true)) { OdtFieldValue right = Comparison(); left = Bool(Truth(left) && Truth(right)); }
        return left;
    }

    private OdtFieldValue Comparison()
    {
        OdtFieldValue left = Additive();
        while (true)
        {
            int compare;
            if (Take("==") || Take("EQ", true)) compare = 0;
            else if (Take("!=") || Take("<>") || Take("NEQ", true)) compare = 1;
            else if (Take("<=") || Take("LEQ", true)) compare = 2;
            else if (Take(">=") || Take("GEQ", true)) compare = 3;
            else if (Take("<") || Take("L", true)) compare = 4;
            else if (Take(">") || Take("G", true)) compare = 5;
            else return left;
            OdtFieldValue right = Additive();
            int order = left.IsText && right.IsText ? string.CompareOrdinal(left.Text, right.Text) : left.AsNumber.CompareTo(right.AsNumber);
            left = Bool(compare switch { 0 => order == 0, 1 => order != 0, 2 => order <= 0, 3 => order >= 0, 4 => order < 0, _ => order > 0 });
        }
    }

    private OdtFieldValue Additive()
    {
        OdtFieldValue left = Multiplicative();
        while (true)
        {
            if (Take("+")) left = OdtFieldValue.Of(left.AsNumber + Multiplicative().AsNumber);
            else if (Take("-")) left = OdtFieldValue.Of(left.AsNumber - Multiplicative().AsNumber);
            else return left;
        }
    }

    private OdtFieldValue Multiplicative()
    {
        OdtFieldValue left = Power();
        while (true)
        {
            if (Take("*") || Take("MUL", true)) left = OdtFieldValue.Of(left.AsNumber * Power().AsNumber);
            else if (Take("/") || Take("DIV", true))
            {
                double divisor = Power().AsNumber;
                if (divisor == 0) throw Bad;                 // SwCalc: division by zero is a calculation error
                left = OdtFieldValue.Of(left.AsNumber / divisor);
            }
            else return left;
        }
    }

    private OdtFieldValue Power()
    {
        OdtFieldValue left = Unary();
        if (Take("^") || Take("POW", true)) return OdtFieldValue.Of(Math.Pow(left.AsNumber, Power().AsNumber));
        return left;
    }

    private OdtFieldValue Unary()
    {
        if (++_depth > 256) throw Bad;
        try { return UnaryInner(); } finally { _depth--; }
    }

    private OdtFieldValue UnaryInner()
    {
        if (Take("-")) return OdtFieldValue.Of(-Unary().AsNumber);
        if (Take("+")) return Unary();
        if (Take("!") || Take("NOT", true)) return Bool(!Truth(Unary()));
        OdtFieldValue value = Primary();
        if (Take("%")) value = OdtFieldValue.Of(value.AsNumber / 100);
        return value;
    }

    private OdtFieldValue Primary()
    {
        SkipBlanks();
        if (_position >= _text.Length) throw Bad;
        char c = _text[_position];
        if (c == '(')
        {
            _position++;
            OdtFieldValue inner = Or();
            if (!Take(")")) throw Bad;
            return inner;
        }
        if (c == '"')
        {
            int end = _text.IndexOf('"', _position + 1);
            if (end < 0) throw Bad;
            string literal = _text.Substring(_position + 1, end - _position - 1);
            _position = end + 1;
            return OdtFieldValue.OfText(literal);
        }
        if (char.IsDigit(c) || c == '.' || c == ',')
        {
            int start = _position;
            while (_position < _text.Length && (char.IsDigit(_text[_position]) || _text[_position] is '.' or ',')) _position++;
            if (_position < _text.Length && _text[_position] is 'e' or 'E' && _position + 1 < _text.Length && (char.IsDigit(_text[_position + 1]) || _text[_position + 1] is '-' or '+'))
            {
                _position += 2;
                while (_position < _text.Length && char.IsDigit(_text[_position])) _position++;
            }
            string number = _text.Substring(start, _position - start).Replace(',', '.');
            if (!double.TryParse(number, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)) throw Bad;
            return OdtFieldValue.Of(parsed);
        }
        if (char.IsLetter(c) || c == '_')
        {
            int start = _position;
            while (_position < _text.Length && (char.IsLetterOrDigit(_text[_position]) || _text[_position] is '_' or '.')) _position++;
            string name = _text.Substring(start, _position - start);
            SkipBlanks();
            if (_position < _text.Length && _text[_position] == '(') return Function(name);
            switch (name.ToUpperInvariant())
            {
                case "PI": return OdtFieldValue.Of(Math.PI);
                case "E": return OdtFieldValue.Of(Math.E);
                case "TRUE": return OdtFieldValue.Of(1);
                case "FALSE": return OdtFieldValue.Of(0);
            }
            return _variable(name) ?? OdtFieldValue.Of(0);   // SwCalc: an unknown name is a new variable with value 0
        }
        throw Bad;
    }

    private OdtFieldValue Function(string name)
    {
        _position++;      // '('
        var arguments = new List<double>();
        if (!Take(")"))
        {
            do arguments.Add(Or().AsNumber);
            while (Take("|") || Take(";"));
            if (!Take(")")) throw Bad;
        }
        double First() => arguments.Count > 0 ? arguments[0] : throw Bad;
        double result = name.ToUpperInvariant() switch
        {
            "ABS" => Math.Abs(First()),
            "SQRT" => First() < 0 ? throw Bad : Math.Sqrt(First()),
            "SIN" => Math.Sin(First()), "COS" => Math.Cos(First()), "TAN" => Math.Tan(First()),
            "ASIN" => Math.Asin(First()), "ACOS" => Math.Acos(First()), "ATAN" => Math.Atan(First()),
            "INT" => Math.Truncate(First()),
            "FRAC" => First() - Math.Truncate(First()),
            "SIGN" => Math.Sign(First()),
            "ROUND" => Math.Round(First(), arguments.Count > 1 ? (int)Math.Max(0, Math.Min(15, arguments[1])) : 0, MidpointRounding.AwayFromZero),
            "POW" => arguments.Count == 2 ? Math.Pow(arguments[0], arguments[1]) : throw Bad,
            "MIN" => arguments.Count > 0 ? arguments.Min() : throw Bad,
            "MAX" => arguments.Count > 0 ? arguments.Max() : throw Bad,
            "SUM" => arguments.Sum(),
            "MEAN" => arguments.Count > 0 ? arguments.Average() : throw Bad,
            _ => throw Bad,
        };
        return OdtFieldValue.Of(result);
    }

    private static bool Truth(OdtFieldValue value) => value.IsText ? value.Text!.Length > 0 : value.Number != 0;
    private static OdtFieldValue Bool(bool value) => OdtFieldValue.Of(value ? 1 : 0);
}

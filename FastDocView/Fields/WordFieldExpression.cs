using System.Globalization;
using System.Text;

namespace FastDocView.Fields;

/// <summary>
/// The formula field's expression (ECMA-376 Part 1 17.16.3 "Formulas"): numbers, + - * / ^ % and the comparisons
/// = &lt;&gt; &lt; &lt;= &gt; &gt;=, parentheses, the functions ABS AND AVERAGE COUNT DEFINED FALSE IF INT MAX MIN MOD NOT OR
/// PRODUCT ROUND SIGN SUM TRUE, table cells (A1, A1:B3, R1C1, LEFT / RIGHT / ABOVE / BELOW) and bookmarks. Checked in Word 16
/// (2026-11-04): -2^2 → 4 (the sign binds tighter than ^), 2+3*4 → 14, 5% → 0.05, 10 > 5 → 1, = 7 MOD 3 is a syntax
/// error (MOD is a function only), = 1e3 a syntax error, a list separator other than "," a syntax error, 2^3^2 → 512 (^ binds
/// right to left). Word computes in binary floating point and shows 14 significant digits (= 1234567890123456 →
/// 1234567890123500, = 10^30 → 1 and 30 zeros); from 10^63 it shows <see cref="NumberTooLarge"/>. Nesting is bounded by Word
/// itself (<see cref="MaximumNesting"/>), so the recursive descent below never goes deeper than that.
/// </summary>
public sealed class WordFieldExpression
{
    /// <summary>Word 16's error texts (Hebrew UI, as written by Word on the dev machine).</summary>
    public const string SyntaxError = "!שגיאת תחביר, ";
    public const string ZeroDivide = "!חילוק באפס";
    public const string UndefinedBookmark = "!סימניה לא מוגדרת, ";
    public const string UnexpectedEnd = "!סוף נוסחה לא צפוי";

    /// <summary>Word 16 (Hebrew UI): "!Formula too complex" - see <see cref="MaximumNesting"/>.</summary>
    public const string TooComplex = "!\u05e0\u05d5\u05e1\u05d7\u05d4 \u05de\u05d5\u05e8\u05db\u05d1\u05ea \u05de\u05d3\u05d9";
    /// <summary>Word 16 (Hebrew UI): "!Number too large to format" - = 10^62 still shows its 63 digits, = 10^63 this text.</summary>
    public const string NumberTooLarge = "!\u05de\u05e1\u05e4\u05e8 \u05d2\u05d3\u05d5\u05dc \u05de\u05db\u05d3\u05d9 \u05e9\u05e0\u05d9\u05ea\u05df \u05d9\u05d4\u05d9\u05d4 \u05dc\u05e2\u05e6\u05d1 \u05d0\u05d5\u05ea\u05d5";

    /// <summary>
    /// Word's own nesting bound, measured in Word 16 through COM (2026-11-04): 27 signs (= ---…-1) or 27 parentheses compute,
    /// 28 give <see cref="TooComplex"/>; = 1+(1+(… with 9 levels computes, 10 does not; ABS(ABS(… and DEFINED(DEFINED(… 6 levels
    /// compute, 7 do not; 2^1^1… 14 times computes, 15 does not. Modelled as nesting units: a sign or "(" costs 1, a binary
    /// operator waiting for its right operand 2, a function call 4; more than 27 units is too complex (the model allows 13
    /// chained ^ where Word allows 14).
    /// </summary>
    public const int MaximumNesting = 27;

    private sealed class FormulaException(string message) : Exception(message);

    private readonly string _text;
    private readonly WordFieldContext _context;
    private int _position;
    private int _nesting;
    private int _decimals;
    private bool _grouped;

    private WordFieldExpression(string text, WordFieldContext context, int nesting) { _text = text; _context = context; _nesting = nesting; }

    /// <summary>The value, or an error text Word would show.</summary>
    public static WordFieldValue Evaluate(string expression, WordFieldContext context) => Evaluate(expression, context, 0);

    private static WordFieldValue Evaluate(string expression, WordFieldContext context, int nesting)
    {
        var parser = new WordFieldExpression(expression, context, nesting);
        try
        {
            parser.SkipSpaces();
            if (parser._position >= parser._text.Length) return WordFieldValue.OfText(UnexpectedEnd);
            double value = parser.Comparison();
            parser.SkipSpaces();
            if (parser._position < parser._text.Length) throw new FormulaException(SyntaxError + parser.NextTokenText());
            return Shown(value, parser._decimals, parser._grouped);
        }
        catch (FormulaException error) { return WordFieldValue.OfText(error.Message); }
    }

    /// <summary>
    /// The value as Word shows it: rounded to 14 significant digits; beyond the decimal type's range the digits as text
    /// (Word 16: = 10^30 → 1 and 30 zeros), from 10^63 <see cref="NumberTooLarge"/>.
    /// </summary>
    private static WordFieldValue Shown(double value, int decimals, bool grouped)
    {
        if (double.IsNaN(value) || double.IsInfinity(value)) return WordFieldValue.OfText(NumberTooLarge);
        if (value == 0) return new WordFieldValue(Number: 0m, Decimals: decimals, Grouped: grouped);
        string scientific = value.ToString("E13", CultureInfo.InvariantCulture);        // -d.ddddddddddddddE+xxx: 14 significant digits
        int exponent = int.Parse(scientific.Substring(scientific.IndexOf('E') + 1), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
        if (exponent >= 63) return WordFieldValue.OfText(NumberTooLarge);
        if (exponent < 28) return new WordFieldValue(Number: decimal.Parse(scientific, NumberStyles.Float, CultureInfo.InvariantCulture), Decimals: decimals, Grouped: grouped);
        string digits = scientific.Substring(value < 0 ? 1 : 0, 15).Replace(".", "").TrimEnd('0');
        return WordFieldValue.OfText((value < 0 ? "-" : "") + digits.PadRight(exponent + 1, '0'));
    }

    /// <summary>A value rounded to the 14 significant digits Word shows (comparisons see what the reader sees).</summary>
    private static double Significant(double value)
        => value == 0 || double.IsNaN(value) || double.IsInfinity(value) ? value : double.Parse(value.ToString("E13", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);

    private void Enter(int units)
    {
        _nesting += units;
        if (_nesting > MaximumNesting) throw new FormulaException(TooComplex);
    }

    private void Leave(int units) => _nesting -= units;

    private void SkipSpaces() { while (_position < _text.Length && char.IsWhiteSpace(_text[_position])) _position++; }

    private string NextTokenText()
    {
        SkipSpaces();
        if (_position >= _text.Length) return "";
        int start = _position;
        if (char.IsLetterOrDigit(_text[start])) { int end = start; while (end < _text.Length && char.IsLetterOrDigit(_text[end])) end++; return _text.Substring(start, end - start); }
        return _text[start].ToString();
    }

    private bool Accept(string symbol)
    {
        SkipSpaces();
        if (string.CompareOrdinal(_text, _position, symbol, 0, symbol.Length) != 0) return false;
        _position += symbol.Length;
        return true;
    }

    private double Comparison()
    {
        double left = Additive();
        SkipSpaces();
        foreach (string op in (string[])["<=", ">=", "<>", "=", "<", ">"])
            if (Accept(op))
            {
                double right = RightOperand(Additive);
                left = Significant(left);
                right = Significant(right);
                bool result = op switch { "<=" => left <= right, ">=" => left >= right, "<>" => left != right, "=" => left == right, "<" => left < right, _ => left > right };
                return result ? 1 : 0;
            }
        return left;
    }

    /// <summary>A binary operator's right operand: it nests while the left operand and the operator wait.</summary>
    private double RightOperand(Func<double> operand)
    {
        Enter(2);
        double value = operand();
        Leave(2);
        return value;
    }

    private double Additive()
    {
        double value = Multiplicative();
        while (true)
        {
            if (Accept("+")) value += RightOperand(Multiplicative);
            else if (Accept("-")) value -= RightOperand(Multiplicative);
            else return value;
        }
    }

    private double Multiplicative()
    {
        double value = Power();
        while (true)
        {
            if (Accept("*")) value *= RightOperand(Power);
            else if (Accept("/"))
            {
                double divisor = RightOperand(Power);
                if (divisor == 0) throw new FormulaException(ZeroDivide);
                value /= divisor;
            }
            else return value;
        }
    }

    /// <summary>^ binds right to left (Word 16: 2^3^2 → 512).</summary>
    private double Power()
    {
        double value = Unary();
        if (!Accept("^")) return value;
        double exponent = RightOperand(Power);
        double result = Math.Pow(value, exponent);
        if (double.IsNaN(result)) throw new FormulaException(SyntaxError + "^");
        return result;
    }

    /// <summary>Signs, then the operand and its % signs. The signs are counted in a loop; each still nests, as in Word.</summary>
    private double Unary()
    {
        int signs = 0;
        bool negative = false;
        while (true)
        {
            if (Accept("-")) negative = !negative;
            else if (!Accept("+")) break;
            Enter(1);
            signs++;
        }
        double value = Primary();
        while (Accept("%")) value /= 100;
        Leave(signs);
        return negative ? -value : value;
    }

    private double Primary()
    {
        SkipSpaces();
        if (_position >= _text.Length) throw new FormulaException(UnexpectedEnd);
        char c = _text[_position];
        if (c == '(')
        {
            _position++;
            Enter(1);
            double value = Comparison();
            if (!Accept(")")) throw new FormulaException(_position >= _text.Length ? UnexpectedEnd : SyntaxError + NextTokenText());
            Leave(1);
            return value;
        }
        if (char.IsDigit(c) || c == '.') return Number();
        if (char.IsLetter(c) || c == '_')
        {
            int start = _position;
            while (_position < _text.Length && (char.IsLetterOrDigit(_text[_position]) || _text[_position] == '_')) _position++;
            string name = _text.Substring(start, _position - start);
            SkipSpaces();
            if (_position < _text.Length && _text[_position] == '(') { _position++; return Function(name.ToUpperInvariant()); }
            if (name.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) return 1;
            if (name.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) return 0;
            List<double>? cells = CellValues(name);
            if (cells is not null) return cells.Count > 0 ? cells[0] : 0;
            return BookmarkValue(name);
        }
        throw new FormulaException(SyntaxError + c);
    }

    private double Number()
    {
        int start = _position;
        int decimals = 0;
        bool afterPoint = false, grouped = false;
        while (_position < _text.Length)
        {
            char c = _text[_position];
            if (char.IsDigit(c)) { if (afterPoint) decimals++; _position++; }
            else if (c == '.' && !afterPoint) { afterPoint = true; _position++; }
            else if (c == ',' && !afterPoint && _position + 3 < _text.Length + 1 && IsGroupOfThree(_position + 1)) { grouped = true; _position++; }
            else break;
        }
        if (_position < _text.Length && char.IsLetter(_text[_position]))
        {
            int end = _position;
            while (end < _text.Length && char.IsLetterOrDigit(_text[end])) end++;
            throw new FormulaException(SyntaxError + _text.Substring(start, end - start).ToUpperInvariant());
        }
        string digits = _text.Substring(start, _position - start).Replace(",", "");
        // a point with no digit (= . or = 1+.) is a syntax error in Word 16; a digit run of any length is a number (Word 16:
        // a 30-digit literal shows its first 14 digits)
        if (!digits.Any(char.IsDigit) || !double.TryParse(digits, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out double value))
            throw new FormulaException(SyntaxError + ".");
        _decimals = Math.Max(_decimals, decimals);
        _grouped |= grouped;
        return value;
    }

    private bool IsGroupOfThree(int at)
        => at + 3 <= _text.Length && char.IsDigit(_text[at]) && char.IsDigit(_text[at + 1]) && char.IsDigit(_text[at + 2])
           && (at + 3 == _text.Length || !char.IsDigit(_text[at + 3]));

    private double BookmarkValue(string name)
    {
        string? text = _context.BookmarkText?.Invoke(name);
        if (text is null) throw new FormulaException(UndefinedBookmark + name.ToUpperInvariant());
        return (double)(LeadingNumber(text) ?? 0);
    }

    /// <summary>The number a cell or bookmark's text starts with (digits, a sign, grouping commas, a point; currency signs skipped).</summary>
    internal static decimal? LeadingNumber(string text)
    {
        var digits = new StringBuilder();
        int i = 0;
        while (i < text.Length && (char.IsWhiteSpace(text[i]) || text[i] is '$' or '€' or '£' or '₪')) i++;
        if (i < text.Length && text[i] is '-' or '+') digits.Append(text[i++]);
        bool any = false;
        for (; i < text.Length; i++)
        {
            char c = text[i];
            if (char.IsDigit(c)) { digits.Append(c); any = true; }
            else if (c == '.') digits.Append(c);
            else if (c == ',' && any) continue;
            else break;
        }
        return any && decimal.TryParse(digits.ToString(), NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out decimal value) ? value : null;
    }

    /// <summary>Arguments of a function: expressions, cell ranges and LEFT / RIGHT / ABOVE / BELOW (each a list of values).</summary>
    private List<double> Arguments()
    {
        var values = new List<double>();
        SkipSpaces();
        if (Accept(")")) return values;
        while (true)
        {
            SkipSpaces();
            int start = _position;
            int end = start;
            while (end < _text.Length && (char.IsLetterOrDigit(_text[end]) || _text[end] == ':')) end++;
            string word = _text.Substring(start, end - start);
            int after = end;
            while (after < _text.Length && char.IsWhiteSpace(_text[after])) after++;
            List<double>? cells = word.Length > 0 && char.IsLetter(word[0]) && (after >= _text.Length || _text[after] is ',' or ')') ? CellValues(word) : null;
            if (cells is not null) { values.AddRange(cells); _position = end; }
            else values.Add(Comparison());
            if (Accept(",")) continue;
            if (Accept(")")) return values;
            throw new FormulaException(_position >= _text.Length ? UnexpectedEnd : SyntaxError + NextTokenText());
        }
    }

    private double Function(string name)
    {
        Enter(4);
        double value = FunctionValue(name);
        Leave(4);
        return value;
    }

    private double FunctionValue(string name)
    {
        if (name == "DEFINED")
        {
            // 1 when the argument evaluates (a known bookmark / cell), else 0 (Word 16: DEFINED(xyz) → 0)
            int depth = 1, start = _position;
            while (_position < _text.Length && depth > 0) { if (_text[_position] == '(') depth++; else if (_text[_position] == ')') depth--; _position++; }
            string inner = _text.Substring(start, Math.Max(0, _position - start - 1));
            // the argument nests inside this call (Word 16: 7 nested DEFINED are too complex)
            WordFieldValue argument = Evaluate(inner, _context, _nesting);
            if (argument.Text == TooComplex) throw new FormulaException(TooComplex);
            return argument.Number is not null || argument.Text is { Length: > 0 } digits && digits[0] != '!' ? 1 : 0;
        }
        List<double> args = Arguments();
        double Arg(int i) => i < args.Count ? args[i] : throw new FormulaException(SyntaxError + ")");
        switch (name)
        {
            case "ABS": return Math.Abs(Arg(0));
            case "AND": return args.Count > 0 && args.All(a => a != 0) ? 1 : 0;
            case "AVERAGE": return args.Count == 0 ? 0 : args.Sum() / args.Count;
            case "COUNT": return args.Count;
            case "FALSE": return 0;
            case "TRUE": return 1;
            case "IF": return Arg(0) != 0 ? Arg(1) : Arg(2);
            case "INT": return Math.Truncate(Arg(0));                               // Word 16: INT(-2.5) → -2
            case "MAX": return args.Count == 0 ? 0 : args.Max();
            case "MIN": return args.Count == 0 ? 0 : args.Min();
            case "MOD":
                if (Arg(1) == 0) throw new FormulaException(ZeroDivide);
                return Arg(0) % Arg(1);
            case "NOT": return Arg(0) == 0 ? 1 : 0;
            case "OR": return args.Any(a => a != 0) ? 1 : 0;
            case "PRODUCT": { double product = 1; foreach (double a in args) product *= a; return product; }
            case "ROUND": return Round(Arg(0), Math.Truncate(Arg(1)));
            case "SIGN": return Math.Sign(Arg(0));
            case "SUM": return args.Sum();
            default: throw new FormulaException(SyntaxError + name);
        }
    }

    /// <summary>
    /// ROUND(x, places), half away from zero on the digits the reader sees (Word 16: ROUND(2.5,0) → 3, ROUND(1.25,1) → 1.3,
    /// ROUND(1.5,40) → 1.5, ROUND(123,-40) → 0). Word 16 gives a syntax error from ROUND(1,32767) up (ROUND(1,99999) and
    /// ROUND(1,2147483647) too) and shows nothing from ROUND(1,-32769) down.
    /// </summary>
    private static double Round(double value, double places)
    {
        if (places >= 32767) throw new FormulaException(SyntaxError + places.ToString("R", CultureInfo.InvariantCulture));
        if (places < -32768) throw new FormulaException("");
        if (double.IsInfinity(value) || Math.Abs(value) >= 7.9e28) return value;
        decimal shown = value == 0 ? 0m : decimal.Parse(value.ToString("E13", CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture);
        if (places >= 0) return (double)Math.Round(shown, (int)Math.Min(places, 28), MidpointRounding.AwayFromZero);
        if (places <= -29) return 0;
        decimal scale = 1;
        for (int i = 0; i < -places; i++) scale *= 10;
        return (double)(Math.Round(shown / scale, 0, MidpointRounding.AwayFromZero) * scale);
    }

    /// <summary>
    /// The values of a cell reference in the field's table, or null when the word is not one: A1, A1:B3, R2C3, LEFT, RIGHT,
    /// ABOVE, BELOW (17.16.3.1 / .2). LEFT / RIGHT / ABOVE / BELOW take the cells in that direction that hold a number,
    /// stopping at the first that does not (Microsoft's "Use a formula in a Word table": "the formula ... stops at the first
    /// blank cell or cell with text" - from the support article, not the standard).
    /// </summary>
    private List<double>? CellValues(string word)
    {
        IWordFieldTable? table = _context.Table;
        string upper = word.ToUpperInvariant();
        if (upper is "LEFT" or "RIGHT" or "ABOVE" or "BELOW")
        {
            if (table is null) throw new FormulaException(SyntaxError + upper);
            var values = new List<double>();
            int row = table.CurrentRow, column = table.CurrentColumn;
            int dRow = upper == "ABOVE" ? -1 : upper == "BELOW" ? 1 : 0, dColumn = upper == "LEFT" ? -1 : upper == "RIGHT" ? 1 : 0;
            for (row += dRow, column += dColumn; row >= 0 && row < table.RowCount && column >= 0 && column < table.ColumnCount(row); row += dRow, column += dColumn)
            {
                if (table.CellText(row, column) is not { } text || LeadingNumber(text) is not { } number) { if (values.Count > 0) break; continue; }
                values.Add((double)number);
            }
            return values;
        }
        string[] ends = upper.Split(':');
        if (ends.Length > 2) return null;
        var first = CellAddress(ends[0]);
        if (first is null) return null;
        var last = ends.Length == 2 ? CellAddress(ends[1]) : first;
        if (last is null) return null;
        if (table is null) return null;
        // the range is the cells that exist inside it (Word 16, 2x2 table: SUM(A1:B9) → the four cells); the bounds are clamped
        // to the table before walking, so a huge range (A1:ZZ2000000000, R1C1:R2000000000C9) costs no more than the table
        var cells = new List<double>();
        int lastRow = Math.Min(Math.Max(first.Value.row, last.Value.row), table.RowCount - 1);
        int firstColumn = Math.Min(first.Value.column, last.Value.column), lastColumn = Math.Max(first.Value.column, last.Value.column);
        for (int r = Math.Min(first.Value.row, last.Value.row); r <= lastRow; r++)
        {
            int rowLastColumn = Math.Min(lastColumn, table.ColumnCount(r) - 1);
            for (int c = firstColumn; c <= rowLastColumn; c++)
                if (table.CellText(r, c) is { } text && LeadingNumber(text) is { } number) cells.Add((double)number);
        }
        return cells;
    }

    /// <summary>A1 (column letters, row number) or RnCn, zero based; null when the word is neither.</summary>
    private static (int row, int column)? CellAddress(string word)
    {
        if (word.Length >= 4 && word[0] == 'R')
        {
            int c = word.IndexOf('C');
            if (c > 1 && int.TryParse(word.Substring(1, c - 1), out int r1) && int.TryParse(word.Substring(c + 1), out int c1) && r1 > 0 && c1 > 0) return (r1 - 1, c1 - 1);
        }
        int i = 0, column = 0;
        while (i < word.Length && word[i] is >= 'A' and <= 'Z') { column = column * 26 + (word[i] - 'A' + 1); i++; }
        if (i == 0 || i > 2 || i == word.Length) return null;
        if (!int.TryParse(word.Substring(i), NumberStyles.None, CultureInfo.InvariantCulture, out int row) || row <= 0) return null;
        return (row - 1, column - 1);
    }
}

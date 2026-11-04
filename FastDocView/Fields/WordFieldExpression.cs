using System.Globalization;
using System.Text;

namespace FastDocView.Fields;

/// <summary>
/// The formula field's expression (ECMA-376 Part 1 17.16.3 "Formulas"): numbers, + - * / ^ % and the comparisons
/// = &lt;&gt; &lt; &lt;= &gt; &gt;=, parentheses, the functions ABS AND AVERAGE COUNT DEFINED FALSE IF INT MAX MIN MOD NOT OR
/// PRODUCT ROUND SIGN SUM TRUE, table cells (A1, A1:B3, R1C1, LEFT / RIGHT / ABOVE / BELOW) and bookmarks. Checked in Word 16
/// (2026-11-04): -2^2 → 4 (the sign binds tighter than ^), 2+3*4 → 14, 5% → 0.05, 10 > 5 → 1, = 7 MOD 3 is a syntax
/// error (MOD is a function only), = 1e3 a syntax error, a list separator other than "," a syntax error.
/// </summary>
public sealed class WordFieldExpression
{
    /// <summary>Word 16's error texts (Hebrew UI, as written by Word on the dev machine).</summary>
    public const string SyntaxError = "!שגיאת תחביר, ";
    public const string ZeroDivide = "!חילוק באפס";
    public const string UndefinedBookmark = "!סימניה לא מוגדרת, ";
    public const string UnexpectedEnd = "!סוף נוסחה לא צפוי";

    private sealed class FormulaException(string message) : Exception(message);

    private readonly string _text;
    private readonly WordFieldContext _context;
    private int _position;
    private int _decimals;
    private bool _grouped;

    private WordFieldExpression(string text, WordFieldContext context) { _text = text; _context = context; }

    /// <summary>The value, or an error text Word would show.</summary>
    public static WordFieldValue Evaluate(string expression, WordFieldContext context)
    {
        var parser = new WordFieldExpression(expression, context);
        try
        {
            parser.SkipSpaces();
            if (parser._position >= parser._text.Length) return WordFieldValue.OfText(UnexpectedEnd);
            decimal value = parser.Comparison();
            parser.SkipSpaces();
            if (parser._position < parser._text.Length) throw new FormulaException(SyntaxError + parser.NextTokenText());
            return new WordFieldValue(Number: value, Decimals: parser._decimals, Grouped: parser._grouped);
        }
        catch (FormulaException error) { return WordFieldValue.OfText(error.Message); }
    }

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

    private decimal Comparison()
    {
        decimal left = Additive();
        SkipSpaces();
        foreach (string op in (string[])["<=", ">=", "<>", "=", "<", ">"])
            if (Accept(op))
            {
                decimal right = Additive();
                bool result = op switch { "<=" => left <= right, ">=" => left >= right, "<>" => left != right, "=" => left == right, "<" => left < right, _ => left > right };
                return result ? 1 : 0;
            }
        return left;
    }

    private decimal Additive()
    {
        decimal value = Multiplicative();
        while (true)
        {
            if (Accept("+")) value += Multiplicative();
            else if (Accept("-")) value -= Multiplicative();
            else return value;
        }
    }

    private decimal Multiplicative()
    {
        decimal value = Power();
        while (true)
        {
            if (Accept("*")) value *= Power();
            else if (Accept("/"))
            {
                decimal divisor = Power();
                if (divisor == 0) throw new FormulaException(ZeroDivide);
                value /= divisor;
            }
            else return value;
        }
    }

    private decimal Power()
    {
        decimal value = Unary();
        while (Accept("^"))
        {
            decimal exponent = Unary();
            double result = Math.Pow((double)value, (double)exponent);
            if (double.IsNaN(result) || double.IsInfinity(result)) throw new FormulaException(SyntaxError + "^");
            value = (decimal)result;
        }
        return value;
    }

    private decimal Unary()
    {
        if (Accept("-")) return -Unary();
        if (Accept("+")) return Unary();
        decimal value = Primary();
        while (Accept("%")) value /= 100;
        return value;
    }

    private decimal Primary()
    {
        SkipSpaces();
        if (_position >= _text.Length) throw new FormulaException(UnexpectedEnd);
        char c = _text[_position];
        if (c == '(')
        {
            _position++;
            decimal value = Comparison();
            if (!Accept(")")) throw new FormulaException(_position >= _text.Length ? UnexpectedEnd : SyntaxError + NextTokenText());
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
            List<decimal>? cells = CellValues(name);
            if (cells is not null) return cells.Count > 0 ? cells[0] : 0;
            return BookmarkValue(name);
        }
        throw new FormulaException(SyntaxError + c);
    }

    private decimal Number()
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
        _decimals = Math.Max(_decimals, decimals);
        _grouped |= grouped;
        return decimal.Parse(digits, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture);
    }

    private bool IsGroupOfThree(int at)
        => at + 3 <= _text.Length && char.IsDigit(_text[at]) && char.IsDigit(_text[at + 1]) && char.IsDigit(_text[at + 2])
           && (at + 3 == _text.Length || !char.IsDigit(_text[at + 3]));

    private decimal BookmarkValue(string name)
    {
        string? text = _context.BookmarkText?.Invoke(name);
        if (text is null) throw new FormulaException(UndefinedBookmark + name.ToUpperInvariant());
        return LeadingNumber(text) ?? 0;
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
    private List<decimal> Arguments()
    {
        var values = new List<decimal>();
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
            List<decimal>? cells = word.Length > 0 && char.IsLetter(word[0]) && (after >= _text.Length || _text[after] is ',' or ')') ? CellValues(word) : null;
            if (cells is not null) { values.AddRange(cells); _position = end; }
            else values.Add(Comparison());
            if (Accept(",")) continue;
            if (Accept(")")) return values;
            throw new FormulaException(_position >= _text.Length ? UnexpectedEnd : SyntaxError + NextTokenText());
        }
    }

    private decimal Function(string name)
    {
        if (name == "DEFINED")
        {
            // 1 when the argument evaluates (a known bookmark / cell), else 0 (Word 16: DEFINED(xyz) → 0)
            int depth = 1, start = _position;
            while (_position < _text.Length && depth > 0) { if (_text[_position] == '(') depth++; else if (_text[_position] == ')') depth--; _position++; }
            string inner = _text.Substring(start, Math.Max(0, _position - start - 1));
            return Evaluate(inner, _context).Number is null ? 0 : 1;
        }
        List<decimal> args = Arguments();
        decimal Arg(int i) => i < args.Count ? args[i] : throw new FormulaException(SyntaxError + ")");
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
            case "PRODUCT": { decimal product = 1; foreach (decimal a in args) product *= a; return product; }
            case "ROUND":
            {
                int places = (int)Arg(1);
                if (places >= 0) return Math.Round(Arg(0), Math.Min(places, 28), MidpointRounding.AwayFromZero);
                decimal scale = (decimal)Math.Pow(10, -places);
                return Math.Round(Arg(0) / scale, 0, MidpointRounding.AwayFromZero) * scale;
            }
            case "SIGN": return Math.Sign(Arg(0));
            case "SUM": return args.Sum();
            default: throw new FormulaException(SyntaxError + name);
        }
    }

    /// <summary>
    /// The values of a cell reference in the field's table, or null when the word is not one: A1, A1:B3, R2C3, LEFT, RIGHT,
    /// ABOVE, BELOW (17.16.3.1 / .2). LEFT / RIGHT / ABOVE / BELOW take the cells in that direction that hold a number,
    /// stopping at the first that does not (Microsoft's "Use a formula in a Word table": "the formula ... stops at the first
    /// blank cell or cell with text" - from the support article, not the standard).
    /// </summary>
    private List<decimal>? CellValues(string word)
    {
        IWordFieldTable? table = _context.Table;
        string upper = word.ToUpperInvariant();
        if (upper is "LEFT" or "RIGHT" or "ABOVE" or "BELOW")
        {
            if (table is null) throw new FormulaException(SyntaxError + upper);
            var values = new List<decimal>();
            int row = table.CurrentRow, column = table.CurrentColumn;
            int dRow = upper == "ABOVE" ? -1 : upper == "BELOW" ? 1 : 0, dColumn = upper == "LEFT" ? -1 : upper == "RIGHT" ? 1 : 0;
            for (row += dRow, column += dColumn; row >= 0 && row < table.RowCount && column >= 0 && column < table.ColumnCount(row); row += dRow, column += dColumn)
            {
                if (table.CellText(row, column) is not { } text || LeadingNumber(text) is not { } number) { if (values.Count > 0) break; continue; }
                values.Add(number);
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
        var cells = new List<decimal>();
        for (int r = Math.Min(first.Value.row, last.Value.row); r <= Math.Max(first.Value.row, last.Value.row); r++)
            for (int c = Math.Min(first.Value.column, last.Value.column); c <= Math.Max(first.Value.column, last.Value.column); c++)
                if (r < table.RowCount && c < table.ColumnCount(r) && table.CellText(r, c) is { } text && LeadingNumber(text) is { } number) cells.Add(number);
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

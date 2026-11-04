using System.Globalization;
using System.Text;

namespace FastDocView.Html;

/// <summary>
/// A custom-shape formula (ODF draw:equation draw:formula, ODF 1.3 part 3 19.171; the OfficeArt presets' equations are
/// turned into the same text first) parsed and evaluated as LibreOffice does
/// (svx/source/customshapes/EnhancedCustomShapeFunctionParser.cxx, master read 2026-11-04, grammar :975-1072):
/// additive = multiplicative (('+' | '-') multiplicative)*; multiplicative = unary (('*' | '/') unary)*; unary = '-' basic | basic;
/// basic = unsigned real (exponent only 'E', :925-935) | keyword (pi left top right bottom xstretch ystretch hasstroke hasfill
/// width height logwidth logheight, matched as prefixes in that order) | '?' name | '$' digits | abs sqrt sin cos tan atan '(' a ')'
/// | min max atan2 '(' a ',' b ')' | if '(' a ',' b ',' c ')' (a &gt; 0, :702-705) | '(' additive ')'; white space is skipped.
/// Angles of sin / cos / tan / atan / atan2 are radians; division by zero follows IEEE (o3tl::div_allow_zero); a formula that
/// does not parse whole is 0 (GetEquationValueAsDouble :912-967), and so is a non-finite result.
/// </summary>
internal abstract class WordCustomShapeFormula
{
    public abstract double Evaluate(WordCustomShapeEngine engine);

    private static readonly string[] Keywords = ["pi", "left", "top", "right", "bottom", "xstretch", "ystretch", "hasstroke", "hasfill", "width", "height", "logwidth", "logheight"];

    /// <summary>The parsed formula; null when it does not parse whole (its value is then 0). Equation references are indices already.</summary>
    public static WordCustomShapeFormula? Parse(string text)
    {
        var parser = new Parser(text);
        WordCustomShapeFormula? result = parser.Additive();
        parser.SkipSpace();
        return result is not null && parser.AtEnd ? result : null;
    }

    /// <summary>
    /// Replaces each "?name" by "?index" of the equation of that name (xmloff ximpcustomshape.cxx:1290-1318; GetEquationName
    /// :171-192 reads [A-Za-z0-9]+; an unknown name is equation 0).
    /// </summary>
    public static string ResolveNames(string formula, IReadOnlyDictionary<string, int> indexOfName)
    {
        if (formula.IndexOf('?') < 0) return formula;
        var result = new StringBuilder(formula.Length);
        for (int i = 0; i < formula.Length; i++)
        {
            result.Append(formula[i]);
            if (formula[i] != '?') continue;
            int end = i + 1;
            while (end < formula.Length && IsNameChar(formula[end])) end++;
            if (end == i + 1) continue;
            string name = formula.Substring(i + 1, end - i - 1);
            result.Append(indexOfName.TryGetValue(name, out int index) ? index : 0);
            i = end - 1;
        }
        return result.ToString();
    }

    internal static bool IsNameChar(char c) => c is >= 'a' and <= 'z' or >= 'A' and <= 'Z' or >= '0' and <= '9';

    private sealed class Constant(double value) : WordCustomShapeFormula
    {
        public override double Evaluate(WordCustomShapeEngine engine) => value;
    }

    private sealed class Keyword(string name) : WordCustomShapeFormula
    {
        public override double Evaluate(WordCustomShapeEngine engine) => engine.Keyword(name);
    }

    private sealed class Adjustment(int index) : WordCustomShapeFormula
    {
        public override double Evaluate(WordCustomShapeEngine engine) => engine.AdjustValue(index);
    }

    private sealed class EquationReference(int index) : WordCustomShapeFormula
    {
        public override double Evaluate(WordCustomShapeEngine engine) => engine.Equation(index);
    }

    private sealed class Unary(string function, WordCustomShapeFormula argument) : WordCustomShapeFormula
    {
        public override double Evaluate(WordCustomShapeEngine engine)
        {
            double a = argument.Evaluate(engine);
            return function switch
            {
                "abs" => Math.Abs(a), "sqrt" => Math.Sqrt(a), "sin" => Math.Sin(a), "cos" => Math.Cos(a), "tan" => Math.Tan(a), "atan" => Math.Atan(a),
                _ => -a,
            };
        }
    }

    private sealed class Binary(char operation, WordCustomShapeFormula first, WordCustomShapeFormula second) : WordCustomShapeFormula
    {
        public override double Evaluate(WordCustomShapeEngine engine)
        {
            double a = first.Evaluate(engine), b = second.Evaluate(engine);
            return operation switch
            {
                '+' => a + b, '-' => a - b, '*' => a * b, '/' => a / b,
                'm' => b < a ? b : a,                       // std::min
                'M' => a < b ? b : a,                       // std::max
                _ => Math.Atan2(a, b),
            };
        }
    }

    private sealed class Conditional(WordCustomShapeFormula condition, WordCustomShapeFormula then, WordCustomShapeFormula otherwise) : WordCustomShapeFormula
    {
        public override double Evaluate(WordCustomShapeEngine engine) => condition.Evaluate(engine) > 0 ? then.Evaluate(engine) : otherwise.Evaluate(engine);
    }

    private sealed class Parser(string text)
    {
        private int _position;

        public bool AtEnd => _position >= text.Length;

        public void SkipSpace() { while (_position < text.Length && char.IsWhiteSpace(text[_position])) _position++; }

        private bool Take(char c)
        {
            SkipSpace();
            if (_position < text.Length && text[_position] == c) { _position++; return true; }
            return false;
        }

        private bool TakeWord(string word)
        {
            SkipSpace();
            if (string.CompareOrdinal(text, _position, word, 0, word.Length) != 0) return false;
            _position += word.Length;
            return true;
        }

        public WordCustomShapeFormula? Additive()
        {
            WordCustomShapeFormula? left = Multiplicative();
            if (left is null) return null;
            while (true)
            {
                int saved = _position;
                if (Take('+')) { WordCustomShapeFormula? right = Multiplicative(); if (right is null) { _position = saved; return left; } left = new Binary('+', left, right); }
                else if (Take('-')) { WordCustomShapeFormula? right = Multiplicative(); if (right is null) { _position = saved; return left; } left = new Binary('-', left, right); }
                else return left;
            }
        }

        private WordCustomShapeFormula? Multiplicative()
        {
            WordCustomShapeFormula? left = UnaryExpression();
            if (left is null) return null;
            while (true)
            {
                int saved = _position;
                if (Take('*')) { WordCustomShapeFormula? right = UnaryExpression(); if (right is null) { _position = saved; return left; } left = new Binary('*', left, right); }
                else if (Take('/')) { WordCustomShapeFormula? right = UnaryExpression(); if (right is null) { _position = saved; return left; } left = new Binary('/', left, right); }
                else return left;
            }
        }

        private WordCustomShapeFormula? UnaryExpression()
        {
            int saved = _position;
            if (Take('-'))
            {
                WordCustomShapeFormula? argument = Basic();
                if (argument is not null) return new Unary("-", argument);
                _position = saved;
            }
            return Basic();
        }

        private WordCustomShapeFormula? Basic()
        {
            SkipSpace();
            int saved = _position;
            if (Number() is { } number) return new Constant(number);
            _position = saved;
            foreach (string keyword in Keywords)
            {
                if (TakeWord(keyword)) return new Keyword(keyword);
                _position = saved;
            }
            if (Take('?'))
            {
                int start = _position;
                while (_position < text.Length && IsNameChar(text[_position])) _position++;
                if (_position > start) return new EquationReference(LeadingInteger(text.Substring(start, _position - start)));
                _position = saved;
            }
            if (Take('$'))
            {
                int start = _position;
                while (_position < text.Length && text[_position] is >= '0' and <= '9') _position++;
                if (_position > start) return new Adjustment(LeadingInteger(text.Substring(start, _position - start)));
                _position = saved;
            }
            foreach (string function in new[] { "abs", "sqrt", "sin", "cos", "tan", "atan" })
            {
                if (TakeWord(function) && Take('(') && Additive() is { } argument && Take(')')) return new Unary(function, argument);
                _position = saved;
            }
            foreach ((string function, char operation) in new[] { ("min", 'm'), ("max", 'M'), ("atan2", 'a') })
            {
                if (TakeWord(function) && Take('(') && Additive() is { } first && Take(',') && Additive() is { } second && Take(')')) return new Binary(operation, first, second);
                _position = saved;
            }
            if (TakeWord("if") && Take('(') && Additive() is { } condition && Take(',') && Additive() is { } then && Take(',') && Additive() is { } otherwise && Take(')'))
                return new Conditional(condition, then, otherwise);
            _position = saved;
            if (Take('(') && Additive() is { } inner && Take(')')) return inner;
            _position = saved;
            return null;
        }

        /// <summary>An unsigned real (boost::spirit ureal_parser: digits, a leading or trailing dot allowed, exponent 'E' with an optional sign).</summary>
        private double? Number()
        {
            int start = _position, digits = 0;
            while (_position < text.Length && char.IsDigit(text[_position])) { _position++; digits++; }
            if (_position < text.Length && text[_position] == '.')
            {
                int dot = _position++;
                int fraction = 0;
                while (_position < text.Length && char.IsDigit(text[_position])) { _position++; fraction++; }
                if (digits == 0 && fraction == 0) { _position = start; return null; }
                digits += fraction;
                _ = dot;
            }
            if (digits == 0) { _position = start; return null; }
            int mantissaEnd = _position;
            if (_position < text.Length && text[_position] == 'E')
            {
                int e = _position++;
                if (_position < text.Length && text[_position] is '+' or '-') _position++;
                int exponentDigits = 0;
                while (_position < text.Length && char.IsDigit(text[_position])) { _position++; exponentDigits++; }
                if (exponentDigits == 0) _position = e;
            }
            _ = mantissaEnd;
            string literal = text.Substring(start, _position - start);
            return double.TryParse(literal, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ? value : 0;
        }

        /// <summary>OUString::toInt32 of the reference text: its leading digits.</summary>
        private static int LeadingInteger(string s)
        {
            long value = 0;
            foreach (char c in s)
            {
                if (c is < '0' or > '9') break;
                value = value * 10 + (c - '0');
                if (value > int.MaxValue) return int.MaxValue;
            }
            return (int)value;
        }
    }

    /// <summary>
    /// An OfficeArt calculation record as the formula LibreOffice makes of it (EnhancedCustomShape2d::GetEquation :88-330,
    /// AppendEnhancedCustomShapeEquationParameter :332-389): special parameters are "?n " (0x400 | n), "$n " (adjust values
    /// 327..336) or left / top / right / bottom (geoLeft..geoBottom 320..323).
    /// </summary>
    public static string OfficeArtEquation(int flags, int p1, int p2, int p3)
    {
        bool s1 = (flags & 0x2000) != 0, s2 = (flags & 0x4000) != 0, s3 = (flags & 0x8000) != 0;
        var e = new StringBuilder();
        void P(int value, bool special)
        {
            if (!special) { e.Append(value.ToString(CultureInfo.InvariantCulture)); return; }
            if ((value & 0x400) != 0) { e.Append('?').Append((value & 0xFF).ToString(CultureInfo.InvariantCulture)).Append(' '); return; }
            switch (value)
            {
                case >= 327 and <= 336: e.Append('$').Append((value - 327).ToString(CultureInfo.InvariantCulture)).Append(' '); break;
                case 320: e.Append("left"); break;
                case 321: e.Append("top"); break;
                case 322: e.Append("right"); break;
                case 323: e.Append("bottom"); break;
            }
        }
        switch (flags & 0xFF)
        {
            case 0: case 14:
            {
                int optimize = (p1 != 0 ? 1 : 0) | (p2 != 0 ? 2 : 0) | (s1 ? 4 : 0) | (s2 ? 8 : 0);
                switch (optimize)
                {
                    case 0: break;
                    case 1: case 4: case 5: P(p1, s1); break;
                    case 2: case 8: case 10: P(p2, s2); break;
                    default: P(p1, s1); e.Append('+'); P(p2, s2); break;
                }
                if (s3 || p3 != 0) { e.Append('-'); P(p3, s3); }
                break;
            }
            case 1:
                P(p1, s1);
                if (s2 || p2 != 1) { e.Append('*'); P(p2, s2); }
                if (s3 || (p3 != 1 && p3 != 0)) { e.Append('/'); P(p3, s3); }
                break;
            case 2: e.Append('('); P(p1, s1); e.Append('+'); P(p2, s2); e.Append(")/2"); break;
            case 3: e.Append("abs("); P(p1, s1); e.Append(')'); break;
            case 4: e.Append("min("); P(p1, s1); e.Append(','); P(p2, s2); e.Append(')'); break;
            case 5: e.Append("max("); P(p1, s1); e.Append(','); P(p2, s2); e.Append(')'); break;
            case 6: e.Append("if("); P(p1, s1); e.Append(','); P(p2, s2); e.Append(','); P(p3, s3); e.Append(')'); break;
            case 7: e.Append("sqrt("); P(p1, s1); e.Append('*'); P(p1, s1); e.Append('+'); P(p2, s2); e.Append('*'); P(p2, s2); e.Append('+'); P(p3, s3); e.Append('*'); P(p3, s3); e.Append(')'); break;
            case 8: e.Append("atan2("); P(p2, s2); e.Append(','); P(p1, s1); e.Append(")/(pi/180)"); break;
            case 9: P(p1, s1); e.Append("*sin("); P(p2, s2); e.Append("*(pi/180))"); break;
            case 10: P(p1, s1); e.Append("*cos("); P(p2, s2); e.Append("*(pi/180))"); break;
            case 11: P(p1, s1); e.Append("*cos(atan2("); P(p3, s3); e.Append(','); P(p2, s2); e.Append("))"); break;
            case 12: P(p1, s1); e.Append("*sin(atan2("); P(p3, s3); e.Append(','); P(p2, s2); e.Append("))"); break;
            case 13: e.Append("sqrt("); P(p1, s1); e.Append(')'); break;
            case 15: P(p3, s3); e.Append("*sqrt(1-("); P(p1, s1); e.Append('/'); P(p2, s2); e.Append(")*("); P(p1, s1); e.Append('/'); P(p2, s2); e.Append("))"); break;
            case 16: P(p1, s1); e.Append("*tan("); P(p2, s2); e.Append(')'); break;
            case 0x80: e.Append("sqrt("); P(p3, s3); e.Append('*'); P(p3, s3); e.Append('-'); P(p1, s1); e.Append('*'); P(p1, s1); e.Append(')'); break;
            case 0x81: e.Append("(cos("); P(p3, s3); e.Append("*(pi/180))*("); P(p1, s1); e.Append("-10800)+sin("); P(p3, s3); e.Append("*(pi/180))*("); P(p2, s2); e.Append("-10800))+10800"); break;
            case 0x82: e.Append("-(sin("); P(p3, s3); e.Append("*(pi/180))*("); P(p1, s1); e.Append("-10800)-cos("); P(p3, s3); e.Append("*(pi/180))*("); P(p2, s2); e.Append("-10800))+10800"); break;
        }
        return e.ToString();
    }
}

/// <summary>A shape's equations, each parsed on first use and its value cached (GetEquationValueAsDouble: bParsed / bReady); a cycle yields 0.</summary>
internal sealed class WordCustomShapeEquations
{
    private readonly string[] _formulas;
    private readonly WordCustomShapeFormula?[] _parsed;
    private readonly bool[] _isParsed, _isReady, _isActive;
    private readonly double[] _values;
    private readonly WordCustomShapeEngine _engine;

    public WordCustomShapeEquations(string[] formulas, WordCustomShapeEngine engine)
    {
        _formulas = formulas;
        _parsed = new WordCustomShapeFormula?[formulas.Length];
        _isParsed = new bool[formulas.Length]; _isReady = new bool[formulas.Length]; _isActive = new bool[formulas.Length];
        _values = new double[formulas.Length];
        _engine = engine;
        engine.Equation = Value;
    }

    public double Value(int index)
    {
        if (index < 0 || index >= _formulas.Length) return 0;
        if (!_isParsed[index]) { _isParsed[index] = true; _parsed[index] = WordCustomShapeFormula.Parse(_formulas[index]); }
        if (_parsed[index] is not { } formula) return 0;
        if (_isReady[index]) return _values[index];
        if (_isActive[index]) return 0;                 // a cycle (LibreOffice would recurse without end)
        _isActive[index] = true;
        double value = formula.Evaluate(_engine);
        _isActive[index] = false;
        if (double.IsNaN(value) || double.IsInfinity(value)) value = 0;
        _values[index] = value;
        _isReady[index] = true;
        return value;
    }
}

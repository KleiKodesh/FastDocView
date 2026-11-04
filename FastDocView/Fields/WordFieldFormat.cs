using System.Globalization;
using System.Text;
using FastDocView.DocxHtml;

namespace FastDocView.Fields;

/// <summary>
/// The general field switches (ECMA-376 Part 1 17.16.4): <c>\*</c> number / text formats, <c>\#</c> numeric pictures and
/// <c>\@</c> date pictures. Behaviour checked in Word 16 on the dev machine (COM automation, 2026-11-04: each result below
/// is what Word's Fields.Update produced) - the harness's WordFieldFormatTest holds those values.
/// </summary>
public static class WordFieldFormat
{
    /// <summary>Word 16's text for a number a format cannot show (Hebrew UI, as Word wrote it on the dev machine: <c>= 393 \* hebrew1</c>).</summary>
    public const string NumberNotRepresentable = "שגיאה! אין אפשרות לייצג את המספר בתבנית שצוינה.";

    /// <summary>
    /// Applies a field's switches to a computed result. A number is formatted by <c>\#</c> first, else by a number format of
    /// <c>\*</c>, else Word's default (<see cref="DefaultNumber"/>); text formats (Upper / Lower / FirstCap / Caps) apply last.
    /// </summary>
    public static string Apply(WordFieldInstruction instruction, WordFieldValue value, WordFieldContext context)
    {
        string text;
        if (value.Date is { } date)
            text = FormatDate(date, instruction.DatePicture ?? value.DefaultDatePicture ?? "M/d/yyyy", context.Culture);
        else if (value.Number is { } number)
        {
            // a \* number format wins over \# (Word 16: = 2 \# "00" \* roman → ii)
            text = instruction.NumericPicture is { } picture ? FormatNumber(number, picture) : DefaultNumber(number, value.Decimals, value.Grouped);
            foreach (string format in instruction.GeneralFormats)
                if (NumberFormat(number, format) is { } formatted) { text = formatted; break; }
        }
        else
        {
            text = value.Text ?? "";
            // a number format on text that reads as a number formats it (Word: QUOTE "21" \* CardText)
            if (instruction.GeneralFormats.Any(IsNumberFormat) && TryParseNumber(text.Trim(), out decimal parsed))
                foreach (string format in instruction.GeneralFormats)
                    if (NumberFormat(parsed, format) is { } formatted) { text = formatted; break; }
            if (instruction.NumericPicture is { } picture && TryParseNumber(text.Trim(), out decimal pictured)) text = FormatNumber(pictured, picture);
            if (instruction.DatePicture is { } datePicture && DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
                text = FormatDate(parsedDate, datePicture, context.Culture);
        }
        foreach (string format in instruction.GeneralFormats) text = CaseFormat(text, format);
        return text;
    }

    internal static bool TryParseNumber(string text, out decimal number)
        => decimal.TryParse(text, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number);

    private static bool IsNumberFormat(string format) => format.ToUpperInvariant() is "ARABIC" or "ARABICDASH" or "ALPHABETIC" or "ROMAN" or "ORDINAL"
        or "ORDTEXT" or "CARDTEXT" or "DOLLARTEXT" or "HEX" or "HEBREW1" or "HEBREW2";

    /// <summary>
    /// A formula's result with no picture: rounded to the most decimals any number in it was written with, at least 2,
    /// trailing zeros dropped (Word 16: = 1/3 → 0.33, = 0.001 → 0.001, = 1.2345/3 → 0.4115, = 3.0 → 3); grouped with commas
    /// when a number in it was (= 1,000 + 1 → 1,001).
    /// </summary>
    public static string DefaultNumber(decimal number, int decimals, bool grouped)
    {
        decimal rounded = Math.Round(number, Math.Max(2, Math.Min(decimals, 28)), MidpointRounding.AwayFromZero);
        if (rounded == 0) rounded = 0;                                               // no "-0"
        string text = rounded.ToString(grouped ? "#,##0.############################" : "0.############################", CultureInfo.InvariantCulture);
        return text;
    }

    /// <summary>A <c>\*</c> number format, or null when the format is not a number format.</summary>
    public static string? NumberFormat(decimal number, string format)
    {
        // Word rounds to a whole number first (= 1.5 \* Arabic → 2, = 2.5 \* roman → iii)
        decimal whole = Math.Round(number, 0, MidpointRounding.AwayFromZero);
        long value = (long)Math.Max(Math.Min(whole, long.MaxValue), long.MinValue);
        switch (format.ToUpperInvariant())
        {
            case "ARABIC": return value.ToString(CultureInfo.InvariantCulture);
            case "ARABICDASH": return "- " + value.ToString(CultureInfo.InvariantCulture) + " -";
            // the first letter's case picks upper or lower (LibreOffice ww8par5.cxx:410-417 GetNumTypeFromName; Word: \* ALPHABETIC / \* alphabetic)
            case "ALPHABETIC": return value <= 0 ? NumberNotRepresentable : RepeatedLetters(value, char.IsUpper(format[0]) ? 'A' : 'a');
            case "ROMAN":
                if (value <= 0) return "";
                string roman = Roman(value);
                return char.IsUpper(format[0]) ? roman : roman.ToLowerInvariant();
            case "ORDINAL": return value < 0 ? NumberNotRepresentable : value.ToString(CultureInfo.InvariantCulture) + OrdinalSuffix(value);
            case "CARDTEXT": return value is < 0 or > 999999 ? NumberNotRepresentable : CardinalText(value);
            case "ORDTEXT": return value is < 0 or > 999999 ? NumberNotRepresentable : OrdinalText(value);
            case "DOLLARTEXT":
            {
                // Word 16: the whole part is checked and spelled as it is; the cents are the fraction rounded, without a carry
                // (999999.999 → "nine hundred ninety-nine thousand nine hundred ninety-nine and 00/100", .994 → 99/100,
                // 1000000 → NumberNotRepresentable)
                if (number < 0 || Math.Floor(number) > 999999) return NumberNotRepresentable;
                long dollars = (long)Math.Floor(number);
                int hundredths = (int)(Math.Round((number - dollars) * 100, 0, MidpointRounding.AwayFromZero) % 100);
                return CardinalText(dollars) + " and " + hundredths.ToString("00", CultureInfo.InvariantCulture) + "/100";
            }
            case "HEX": return value < 0 ? NumberNotRepresentable : value.ToString("X", CultureInfo.InvariantCulture);
            case "HEBREW1": return value is <= 0 or > 392 ? NumberNotRepresentable : DocxNumberingDefinitions.HebrewNumeral((int)value);
            case "HEBREW2": return value <= 0 ? NumberNotRepresentable : HebrewAlphabetField(value);
            default: return null;
        }
    }

    /// <summary>Word's field hebrew2: one ת per full 22, then the letter (Word 16: 22 → ת, 23 → תא, 44 → תת, 45 → תתא) - unlike a list's hebrew2, which repeats the letter.</summary>
    private static string HebrewAlphabetField(long value)
    {
        const string alphabet = "אבגדהוזחטיכלמנסעפצקרשת";
        long full = (value - 1) / alphabet.Length;
        int rest = (int)((value - 1) % alphabet.Length);
        if (full > 10_000) return NumberNotRepresentable;
        var text = new StringBuilder();
        text.Append('ת', (int)full);
        if (rest != alphabet.Length - 1 || full == 0) text.Append(alphabet[rest]);
        else text.Append('ת');
        // 44 = 2 x 22: "תת" (one ת for the first 22, then the 22nd letter)
        return text.ToString();
    }

    /// <summary>A, B, … Z, AA, BB, … (Word 16: 28 → BB, 52 → zz, 53 → aaa).</summary>
    private static string RepeatedLetters(long value, char first)
    {
        long repeat = (value - 1) / 26 + 1;
        if (repeat > 10_000) return NumberNotRepresentable;
        return new string((char)(first + (value - 1) % 26), (int)repeat);
    }

    /// <summary>Roman numerals; Word continues with M past 3999 (Word 16: 4000 → mmmm).</summary>
    private static string Roman(long value)
    {
        if (value / 1000 > 10_000) return NumberNotRepresentable;
        var roman = new StringBuilder();
        roman.Append('M', (int)(value / 1000));
        value %= 1000;
        (int number, string numeral)[] table = [(900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")];
        foreach ((int number, string numeral) in table) while (value >= number) { roman.Append(numeral); value -= number; }
        return roman.ToString();
    }

    private static string OrdinalSuffix(long value)
    {
        long lastTwo = value % 100;
        if (lastTwo is 11 or 12 or 13) return "th";
        return (value % 10) switch { 1 => "st", 2 => "nd", 3 => "rd", _ => "th" };
    }

    private static readonly string[] Ones = ["zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen"];
    private static readonly string[] Tens = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"];

    /// <summary>English number words as Word writes them (Word 16: 21 → twenty-one, 1234 → one thousand two hundred thirty-four, 0 → zero; no "and").</summary>
    public static string CardinalText(long value)
    {
        if (value == 0) return "zero";
        var words = new List<string>();
        if (value >= 1000) { words.Add(UnderThousand(value / 1000)); words.Add("thousand"); value %= 1000; }
        if (value > 0) words.Add(UnderThousand(value));
        return string.Join(" ", words);
    }

    private static string UnderThousand(long value)
    {
        var words = new List<string>();
        if (value >= 100) { words.Add(Ones[value / 100]); words.Add("hundred"); value %= 100; }
        if (value >= 20) words.Add(Tens[value / 10] + (value % 10 > 0 ? "-" + Ones[value % 10] : ""));
        else if (value > 0) words.Add(Ones[value]);
        return string.Join(" ", words);
    }

    /// <summary>English ordinal words (Word 16: 21 → twenty-first, 101 → one hundred first, 112 → one hundred twelfth).</summary>
    public static string OrdinalText(long value)
    {
        string cardinal = CardinalText(value);
        int split = Math.Max(cardinal.LastIndexOf(' '), cardinal.LastIndexOf('-')) + 1;
        string last = cardinal.Substring(split);
        string ordinal = last switch
        {
            "zero" => "zeroth", "one" => "first", "two" => "second", "three" => "third", "five" => "fifth", "eight" => "eighth",
            "nine" => "ninth", "twelve" => "twelfth",
            _ when last.EndsWith("y", StringComparison.Ordinal) => last.Substring(0, last.Length - 1) + "ieth",
            _ => last + "th",
        };
        return cardinal.Substring(0, split) + ordinal;
    }

    /// <summary>Upper / Lower / FirstCap / Caps (Word 16: QUOTE "abc def" \* FirstCap → Abc def, \* Caps → Abc Def); any other format leaves the text.</summary>
    public static string CaseFormat(string text, string format)
    {
        switch (format.ToUpperInvariant())
        {
            case "UPPER": return text.ToUpperInvariant();
            case "LOWER": return text.ToLowerInvariant();
            case "FIRSTCAP":
                for (int i = 0; i < text.Length; i++)
                    if (char.IsLetter(text[i])) return text.Substring(0, i) + char.ToUpperInvariant(text[i]) + text.Substring(i + 1);
                return text;
            case "CAPS":
            {
                var result = new StringBuilder(text.Length);
                bool atWordStart = true;
                foreach (char c in text)
                {
                    result.Append(atWordStart && char.IsLetter(c) ? char.ToUpperInvariant(c) : c);
                    atWordStart = !char.IsLetterOrDigit(c);
                }
                return result.ToString();
            }
            default: return text;
        }
    }

    /// <summary>
    /// A numeric picture (17.16.4.2): <c>0</c> a digit or zero, <c>#</c> a digit or a space, <c>x</c> a digit with the digits
    /// left of it dropped, <c>.</c> the decimal point, <c>,</c> grouping, <c>-</c> / <c>+</c> sign places, <c>'text'</c>
    /// literal, <c>;</c> positive;negative;zero sections; anything else is literal. Rounds half away from zero (Word 16:
    /// 2.5 → 3, -2.5 → -3, 1.005 \# "0.00" → 1.01, -0.4 \# "0" → 0, 0.5 \# "#.##" → " .5 ", 12 \# "x##" → " 12",
    /// 1234.5678 \# ",0.0" → 1,234.6, -5 \# "0.0;(0.0)" → (5.0), 0 \# "...;'zero'" → zero).
    /// </summary>
    public static string FormatNumber(decimal number, string picture)
    {
        List<string> sections = SplitSections(picture);
        string section = sections[0];
        bool isNegativeSection = false;
        if (number < 0 && sections.Count > 1) { section = sections[1]; isNegativeSection = true; }
        int decimals = FractionPlaces(section);
        decimal rounded = Math.Round(Math.Abs(number), Math.Min(decimals, 28), MidpointRounding.AwayFromZero);
        if (rounded == 0 && sections.Count > 2) { section = sections[2]; isNegativeSection = false; decimals = FractionPlaces(section); rounded = 0; }
        bool negative = number < 0 && rounded != 0 && !isNegativeSection;

        // the picture's digit places
        int integerPlaces = 0, xPosition = -1;
        bool grouping = false, seenPoint = false, hasSignPlace = false;
        for (int i = 0; i < section.Length; i++)
        {
            char c = section[i];
            if (c == '\'') { int close = section.IndexOf('\'', i + 1); i = close < 0 ? section.Length : close; continue; }
            if (c == '.') seenPoint = true;
            else if (!seenPoint && c is '0' or '#' or 'x') { if (c == 'x') xPosition = integerPlaces; integerPlaces++; }
            else if (!seenPoint && c == ',') grouping = true;
            else if (c is '-' or '+') hasSignPlace = true;
        }
        string digits = rounded.ToString("0." + new string('0', decimals), CultureInfo.InvariantCulture);
        string integerDigits = digits.Split('.')[0];
        string fractionDigits = digits.Contains('.') ? digits.Split('.')[1] : "";
        if (integerDigits == "0") integerDigits = "";
        // integer places filled right to left; extra digits go before the first place unless an x drops them
        char[] places = new char[integerPlaces];
        string overflow = "";
        {
            int digit = integerDigits.Length - 1;
            int placeIndex = integerPlaces - 1;
            string kinds = PlaceKinds(section, integer: true);
            for (; placeIndex >= 0; placeIndex--, digit--)
                places[placeIndex] = digit >= 0 ? integerDigits[digit] : kinds[placeIndex] == '0' ? '0' : ' ';
            if (digit >= 0 && xPosition < 0) overflow = integerDigits.Substring(0, digit + 1);
        }
        string integerText = overflow + new string(places);
        if (grouping) integerText = Group(integerText);

        var text = new StringBuilder();
        if (negative && !hasSignPlace) text.Append('-');
        bool integerWritten = false;
        int fractionIndex = 0;
        string fractionKinds = PlaceKinds(section, integer: false);
        seenPoint = false;
        for (int i = 0; i < section.Length; i++)
        {
            char c = section[i];
            if (c == '\'')
            {
                int close = section.IndexOf('\'', i + 1);
                if (close < 0) close = section.Length;
                text.Append(section, i + 1, close - i - 1);
                i = close;
                continue;
            }
            if (c == '.') { seenPoint = true; text.Append('.'); continue; }
            if (!seenPoint && c is '0' or '#' or 'x' or ',')
            {
                if (!integerWritten) { text.Append(integerText); integerWritten = true; }
                continue;
            }
            if (seenPoint && c is '0' or '#' or 'x')
            {
                char digit = fractionIndex < fractionDigits.Length ? fractionDigits[fractionIndex] : '0';
                // a # place shows a space for a zero no later digit needs
                bool trailingZero = fractionDigits.Substring(Math.Min(fractionIndex, fractionDigits.Length)).All(d => d == '0');
                text.Append(fractionKinds[fractionIndex] == '0' || !trailingZero ? digit : ' ');
                fractionIndex++;
                continue;
            }
            // in the negative section the sign is written as it stands (Word 16: -5 \# "+0;-0" → -5)
            if (c == '-' && !isNegativeSection) { text.Append(negative ? '-' : ' '); continue; }
            if (c == '+' && !isNegativeSection) { text.Append(negative ? '-' : rounded == 0 ? ' ' : '+'); continue; }
            text.Append(c);
        }
        return text.ToString();
    }

    private static List<string> SplitSections(string picture)
    {
        var sections = new List<string>();
        int start = 0;
        bool quoted = false;
        for (int i = 0; i < picture.Length; i++)
        {
            if (picture[i] == '\'') quoted = !quoted;
            else if (picture[i] == ';' && !quoted) { sections.Add(picture.Substring(start, i - start)); start = i + 1; }
        }
        sections.Add(picture.Substring(start));
        return sections;
    }

    private static int FractionPlaces(string section) => PlaceKinds(section, integer: false).Length;

    /// <summary>The digit places of the integer or fraction part, in order ('0', '#' or 'x').</summary>
    private static string PlaceKinds(string section, bool integer)
    {
        var kinds = new StringBuilder();
        bool seenPoint = false;
        for (int i = 0; i < section.Length; i++)
        {
            char c = section[i];
            if (c == '\'') { int close = section.IndexOf('\'', i + 1); i = close < 0 ? section.Length : close; continue; }
            if (c == '.') { seenPoint = true; continue; }
            if (c is '0' or '#' or 'x' && seenPoint != integer) kinds.Append(c);
        }
        return kinds.ToString();
    }

    /// <summary>Commas between every three digits of the digit run (leading spaces of # places stay spaces).</summary>
    private static string Group(string integerText)
    {
        int firstDigit = 0;
        while (firstDigit < integerText.Length && integerText[firstDigit] == ' ') firstDigit++;
        string digits = integerText.Substring(firstDigit);
        var grouped = new StringBuilder();
        for (int i = 0; i < digits.Length; i++)
        {
            if (i > 0 && (digits.Length - i) % 3 == 0) grouped.Append(',');
            grouped.Append(digits[i]);
        }
        return integerText.Substring(0, firstDigit) + grouped;
    }

    /// <summary>
    /// A date picture (17.16.4.1): M / MM / MMM / MMMM month, d / dd / ddd / dddd day, yy / yyyy year, h / hh 12-hour,
    /// H / HH 24-hour, m / mm minutes, s / ss seconds, am/pm / AM/PM, 'text' literal, anything else literal. Day and month
    /// names come from <paramref name="culture"/> - Word takes them from the system's regional settings (Word 16 on a
    /// Hebrew system wrote "יום שלישי, מרץ 5, 2024" for "dddd, MMMM d, yyyy" whatever the text's language).
    /// </summary>
    public static string FormatDate(DateTime date, string picture, CultureInfo culture)
    {
        DateTimeFormatInfo names = culture.DateTimeFormat;
        var text = new StringBuilder();
        for (int i = 0; i < picture.Length;)
        {
            char c = picture[i];
            if (c == '\'')
            {
                int close = picture.IndexOf('\'', i + 1);
                if (close < 0) close = picture.Length;
                text.Append(picture, i + 1, close - i - 1);
                i = close + 1;
                continue;
            }
            if (string.Compare(picture, i, "am/pm", 0, 5, StringComparison.OrdinalIgnoreCase) == 0)
            {
                string designator = date.Hour < 12 ? names.AMDesignator : names.PMDesignator;
                if (designator.Length == 0) designator = date.Hour < 12 ? "AM" : "PM";
                text.Append(char.IsUpper(c) ? designator.ToUpperInvariant() : designator.ToLowerInvariant());
                i += 5;
                continue;
            }
            int run = 1;
            while (i + run < picture.Length && picture[i + run] == c) run++;
            switch (c)
            {
                case 'M':
                    text.Append(run switch { 1 => date.Month.ToString(CultureInfo.InvariantCulture), 2 => date.Month.ToString("00", CultureInfo.InvariantCulture),
                        3 => names.GetAbbreviatedMonthName(date.Month), _ => names.GetMonthName(date.Month) });
                    break;
                case 'd' or 'D':
                    text.Append(run switch { 1 => date.Day.ToString(CultureInfo.InvariantCulture), 2 => date.Day.ToString("00", CultureInfo.InvariantCulture),
                        3 => names.GetAbbreviatedDayName(date.DayOfWeek), _ => names.GetDayName(date.DayOfWeek) });
                    break;
                case 'y' or 'Y':
                    text.Append(run <= 2 ? (date.Year % 100).ToString("00", CultureInfo.InvariantCulture) : date.Year.ToString(CultureInfo.InvariantCulture));
                    break;
                case 'h':
                {
                    int hour = date.Hour % 12 == 0 ? 12 : date.Hour % 12;
                    text.Append(run == 1 ? hour.ToString(CultureInfo.InvariantCulture) : hour.ToString("00", CultureInfo.InvariantCulture));
                    break;
                }
                case 'H': text.Append(run == 1 ? date.Hour.ToString(CultureInfo.InvariantCulture) : date.Hour.ToString("00", CultureInfo.InvariantCulture)); break;
                case 'm': text.Append(run == 1 ? date.Minute.ToString(CultureInfo.InvariantCulture) : date.Minute.ToString("00", CultureInfo.InvariantCulture)); break;
                case 's' or 'S': text.Append(run == 1 ? date.Second.ToString(CultureInfo.InvariantCulture) : date.Second.ToString("00", CultureInfo.InvariantCulture)); break;
                default: text.Append(c, run); break;
            }
            i += run;
        }
        return text.ToString();
    }
}

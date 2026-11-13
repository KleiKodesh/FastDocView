using System.Globalization;
using System.Text;

namespace FastDocView.Html;

/// <summary>
/// LibreOffice's list numbering types (offapi com/sun/star/style/NumberingType.idl), in the order of
/// i18npool defaultnumberingprovider.cxx:1078 aSupportedTypes - the order its identifier lookup walks.
/// </summary>
public enum ListNumberingType
{
    CharsUpperLetter, CharsLowerLetter, RomanUpper, RomanLower, Arabic, NumberNone, CharSpecial, PageDescriptor, Bitmap,
    SymbolChicago, TextNumber, TextCardinal, TextOrdinal, CharsUpperLetterN, CharsLowerLetterN, NativeNumbering, FullwidthArabic,
    CircleNumber, NumberLowerZh, NumberUpperZh, NumberUpperZhTw, TianGanZh, DiZiZh, NumberTraditionalJa, AiuFullwidthJa,
    AiuHalfwidthJa, IrohaFullwidthJa, IrohaHalfwidthJa, NumberUpperKo, NumberHangulKo, HangulJamoKo, HangulSyllableKo,
    HangulCircledJamoKo, HangulCircledSyllableKo, NumberLegalKo, NumberDigitalKo, NumberDigital2Ko, CharsArabic, CharsArabicAbjad,
    NumberArabicIndic, NumberEastArabicIndic, NumberIndicDevanagari, CharsThai, CharsHebrew, NumberHebrew, CharsNepali, CharsKhmer,
    CharsLao, CharsMyanmar, CharsTibetan, CharsCyrillicUpperLetterBg, CharsCyrillicLowerLetterBg, CharsCyrillicUpperLetterNBg,
    CharsCyrillicLowerLetterNBg, CharsCyrillicUpperLetterRu, CharsCyrillicLowerLetterRu, CharsCyrillicUpperLetterNRu,
    CharsCyrillicLowerLetterNRu, CharsCyrillicUpperLetterSr, CharsCyrillicLowerLetterSr, CharsCyrillicUpperLetterNSr,
    CharsCyrillicLowerLetterNSr, CharsCyrillicUpperLetterUk, CharsCyrillicLowerLetterUk, CharsCyrillicUpperLetterNUk,
    CharsCyrillicLowerLetterNUk, CharsPersian, CharsPersianWord, SzekelyRovas, CharsGreekUpperLetter, CharsGreekLowerLetter,
    ArabicZero, ArabicZero3, ArabicZero4, ArabicZero5,
}

/// <summary>
/// List labels exactly as LibreOffice draws them: i18npool defaultnumberingprovider.cxx makeNumberingString (:642) with the
/// native-number conversions of nativenumbersupplier.cxx (NATNUM1/3/5/8/9/11: AsciiToNative :187, AsciiToNativeChar :85,
/// Hebrew :1055) and, for the spelled-out types (NATNUM12, libnumbertext data/en.sor, zh.sor, hu_Hung.sor), the rules those
/// Soros programs apply to list values. A value below 1 prints nothing, as the provider throws there
/// (editeng numitem.cxx:122 GetNumStr: only ARABIC prints 0). ODF names a type by style:num-format: "1 a A i I" and ""
/// (xmloff xmluconv.cxx:652 convertNumFormat, style:num-letter-sync="true" turning a / A into the repeating aa, bb), any other
/// value is the provider's identifier (makeNumberingIdentifier :1160: the fixed symbol, else the labels for 1, 2, 3 + ", ...");
/// an unknown identifier is ARABIC.
/// </summary>
public static class ListNumberingTypes
{
    /// <summary>The type an ODF style:num-format (+ style:num-letter-sync) names; null for "" (no number).</summary>
    public static ListNumberingType? FromOdfNumberFormat(string? format, bool letterSync)
    {
        if (string.IsNullOrEmpty(format)) return null;
        if (format!.Length == 1)
            switch (format[0])
            {
                case '1': return ListNumberingType.Arabic;
                case 'a': return letterSync ? ListNumberingType.CharsLowerLetterN : ListNumberingType.CharsLowerLetter;
                case 'A': return letterSync ? ListNumberingType.CharsUpperLetterN : ListNumberingType.CharsUpperLetter;
                case 'i': return ListNumberingType.RomanLower;
                case 'I': return ListNumberingType.RomanUpper;
            }
        return (_identifiers ??= BuildIdentifiers()).TryGetValue(format, out ListNumberingType type) ? type : ListNumberingType.Arabic;
    }

    /// <summary>The identifier LibreOffice writes as style:num-format for a type (xmluconv.cxx:713 convertNumFormat + provider :1261).</summary>
    public static string IdentifierOf(ListNumberingType type) => type switch
    {
        ListNumberingType.CharsLowerLetterN => "a",
        ListNumberingType.CharsUpperLetterN => "A",
        ListNumberingType.NumberNone => "",
        _ => MakeIdentifier(type),
    };

    // built on first use: the identifiers are generated from the tables below, which a field initializer here would precede
    private static Dictionary<string, ListNumberingType>? _identifiers;

    private static Dictionary<string, ListNumberingType> BuildIdentifiers()
    {
        var identifiers = new Dictionary<string, ListNumberingType>(StringComparer.Ordinal);
        for (var type = ListNumberingType.CharsUpperLetter; type <= ListNumberingType.ArabicZero5; type++)
        {
            string identifier = MakeIdentifier(type);
            if (!identifiers.ContainsKey(identifier)) identifiers[identifier] = type;   // first in table order wins (:1238)
        }
        return identifiers;
    }

    // defaultnumberingprovider.cxx:1078-1157 fixed symbols; the rest are generated (:1167-1206)
    private static string MakeIdentifier(ListNumberingType type) => type switch
    {
        ListNumberingType.CharsUpperLetter => "A",
        ListNumberingType.CharsLowerLetter => "a",
        ListNumberingType.RomanUpper => "I",
        ListNumberingType.RomanLower => "i",
        ListNumberingType.Arabic => "1",
        ListNumberingType.NumberNone => "''",
        ListNumberingType.CharSpecial => "Bullet",
        ListNumberingType.PageDescriptor => "Page",
        ListNumberingType.Bitmap => "Bitmap",
        ListNumberingType.SymbolChicago => "*, †, ‡, §, **, ††, ...",
        ListNumberingType.TextNumber => "1st",
        ListNumberingType.TextCardinal => "One",
        ListNumberingType.TextOrdinal => "First",
        ListNumberingType.CharsUpperLetterN => "AAA",
        ListNumberingType.CharsLowerLetterN => "aaa",
        ListNumberingType.NativeNumbering => "Native Numbering",
        ListNumberingType.NumberLowerZh => "一, 二, 三, ...",
        ListNumberingType.NumberArabicIndic => "١, ٢, ٣, ...",
        ListNumberingType.NumberEastArabicIndic => "۱, ۲, ۳, ...",
        ListNumberingType.NumberIndicDevanagari => "१, २, ३, ...",
        ListNumberingType.NumberHebrew => "א, י, ק, ...",
        ListNumberingType.CharsCyrillicUpperLetterBg => "А, Б, .., Аа, Аб, ... (bg)",
        ListNumberingType.CharsCyrillicLowerLetterBg => "а, б, .., аа, аб, ... (bg)",
        ListNumberingType.CharsCyrillicUpperLetterNBg => "А, Б, .., Аа, Бб, ... (bg)",
        ListNumberingType.CharsCyrillicLowerLetterNBg => "а, б, .., аа, бб, ... (bg)",
        ListNumberingType.CharsCyrillicUpperLetterRu => "А, Б, .., Аа, Аб, ... (ru)",
        ListNumberingType.CharsCyrillicLowerLetterRu => "а, б, .., аа, аб, ... (ru)",
        ListNumberingType.CharsCyrillicUpperLetterNRu => "А, Б, .., Аа, Бб, ... (ru)",
        ListNumberingType.CharsCyrillicLowerLetterNRu => "а, б, .., аа, бб, ... (ru)",
        ListNumberingType.CharsCyrillicUpperLetterSr => "А, Б, .., Аа, Аб, ... (sr)",
        ListNumberingType.CharsCyrillicLowerLetterSr => "а, б, .., аа, аб, ... (sr)",
        ListNumberingType.CharsCyrillicUpperLetterNSr => "А, Б, .., Аа, Бб, ... (sr)",
        ListNumberingType.CharsCyrillicLowerLetterNSr => "а, б, .., аа, бб, ... (sr)",
        ListNumberingType.CharsCyrillicUpperLetterUk => "А, Б, .., Аа, Аб, ... (uk)",
        ListNumberingType.CharsCyrillicLowerLetterUk => "а, б, .., аа, аб, ... (uk)",
        ListNumberingType.CharsCyrillicUpperLetterNUk => "А, Б, .., Аа, Бб, ... (uk)",
        ListNumberingType.CharsCyrillicLowerLetterNUk => "а, б, .., аа, бб, ... (uk)",
        ListNumberingType.CharsGreekUpperLetter => "Α, Β, ... (gr)",
        ListNumberingType.CharsGreekLowerLetter => "α, β, ... (gr)",
        ListNumberingType.ArabicZero => "01, 02, 03, ...",
        ListNumberingType.ArabicZero3 => "001, 002, 003, ...",
        ListNumberingType.ArabicZero4 => "0001, 0002, 0003, ...",
        ListNumberingType.ArabicZero5 => "00001, 00002, 00003, ...",
        // generated with the "en" locale; :1189-1196 make two duplicates unique
        _ => Format(1, type) + ", " + Format(2, type) + ", " + Format(3, type) + ", ..." + type switch
        {
            ListNumberingType.NumberDigitalKo => " (ko-x-digital)",
            ListNumberingType.NumberDigital2Ko => " (ko)",
            _ => "",
        },
    };

    /// <summary>The label for <paramref name="value"/> (prefix and suffix not included), as makeNumberingString with the "en" locale.</summary>
    public static string Format(int value, ListNumberingType type)
    {
        if (type == ListNumberingType.Arabic) return value.ToString(CultureInfo.InvariantCulture);   // numitem.cxx:138 0 allowed
        if (value <= 0) return "";
        string digits = value.ToString(CultureInfo.InvariantCulture);
        return type switch
        {
            ListNumberingType.CharsUpperLetter => Bijective(value, UpperLetters),
            ListNumberingType.CharsLowerLetter => Bijective(value, LowerLetters),
            ListNumberingType.TextNumber => EnglishOrdinalNumber(value),
            ListNumberingType.TextCardinal => Capitalize(EnglishCardinal(digits, top: true)),
            ListNumberingType.TextOrdinal => Capitalize(EnglishOrdinal(EnglishCardinal(digits, top: true))),
            ListNumberingType.RomanUpper => Roman(value),
            ListNumberingType.RomanLower => Roman(value).ToLowerInvariant(),
            ListNumberingType.NumberNone or ListNumberingType.CharSpecial or ListNumberingType.PageDescriptor or ListNumberingType.Bitmap => "",
            ListNumberingType.CharsUpperLetterN => Repeated(value, UpperLetters),
            ListNumberingType.CharsLowerLetterN => Repeated(value, LowerLetters),
            ListNumberingType.NativeNumbering => digits,           // NATNUM1 for "en": no native digits (getLanguageNumber -1)
            ListNumberingType.FullwidthArabic => digits,            // NATNUM3 for "en": getLanguageNumber -1 returns the ASCII digits (:785)
            ListNumberingType.NumberLowerZh => ChineseText(digits, top: true),
            ListNumberingType.NumberUpperZh => AsciiToNative(digits, NatNum5SimplifiedChinese),
            ListNumberingType.NumberUpperZhTw => AsciiToNative(digits, NatNum5TraditionalChinese),
            ListNumberingType.NumberTraditionalJa => AsciiToNative(digits, NatNum8Japanese),
            ListNumberingType.NumberUpperKo => AsciiToNative(digits, NatNum8Korean),
            ListNumberingType.NumberHangulKo => AsciiToNative(digits, NatNum11Korean),
            ListNumberingType.NumberDigitalKo => DigitByDigit(digits, HangulDigits),
            ListNumberingType.NumberDigital2Ko => DigitByDigit(digits, KoreanLowerDigits),
            ListNumberingType.NumberLegalKo => value < 100 ? KoreanLegalWord(value) : AsciiToNative(digits, NatNum11Korean),
            ListNumberingType.CircleNumber => Table(value, CircledNumbers, recycle: false),
            ListNumberingType.TianGanZh => Table(value, TianGan, recycle: false),
            ListNumberingType.DiZiZh => Table(value, DiZi, recycle: false),
            ListNumberingType.AiuFullwidthJa => Table(value, WordListNumberScripts.AiueoFullWidth, recycle: true),
            ListNumberingType.AiuHalfwidthJa => Table(value, WordListNumberScripts.AiueoHalfWidth, recycle: true),
            ListNumberingType.IrohaFullwidthJa => Table(value, WordListNumberScripts.IrohaFullWidth, recycle: true),
            ListNumberingType.IrohaHalfwidthJa => Table(value, WordListNumberScripts.IrohaHalfWidth, recycle: true),
            ListNumberingType.HangulJamoKo => Table(value, WordListNumberScripts.HangulJamo, recycle: true),
            ListNumberingType.HangulSyllableKo => Table(value, WordListNumberScripts.HangulSyllables, recycle: true),
            ListNumberingType.HangulCircledJamoKo => Table(value, HangulCircledJamo, recycle: true),
            ListNumberingType.HangulCircledSyllableKo => Table(value, HangulCircledSyllables, recycle: true),
            ListNumberingType.CharsArabic => Bijective(value, ArabicLetters),
            ListNumberingType.CharsArabicAbjad => Bijective(value, ArabicAbjadLetters),
            ListNumberingType.NumberArabicIndic => DigitByDigit(digits, ArabicIndicDigits),
            ListNumberingType.NumberEastArabicIndic => DigitByDigit(digits, EastArabicIndicDigits),
            ListNumberingType.NumberIndicDevanagari => DigitByDigit(digits, DevanagariDigits),
            ListNumberingType.CharsThai => Bijective(value, ThaiLetters),
            ListNumberingType.CharsHebrew => Bijective(value, HebrewLetters),
            ListNumberingType.NumberHebrew => HebrewNumber(value),
            ListNumberingType.CharsNepali => Bijective(value, NepaliLetters),
            ListNumberingType.CharsKhmer => Bijective(value, KhmerLetters),
            ListNumberingType.CharsLao => Bijective(value, LaoLetters),
            ListNumberingType.CharsMyanmar => Bijective(value, MyanmarLetters),
            ListNumberingType.CharsTibetan => Bijective(value, TibetanLetters),
            ListNumberingType.CharsCyrillicUpperLetterBg => BijectiveCapital(value, CyrillicUpperBg, CyrillicLowerBg),
            ListNumberingType.CharsCyrillicLowerLetterBg => Bijective(value, CyrillicLowerBg),
            ListNumberingType.CharsCyrillicUpperLetterNBg => RepeatedCapital(value, CyrillicUpperBg, CyrillicLowerBg),
            ListNumberingType.CharsCyrillicLowerLetterNBg => Repeated(value, CyrillicLowerBg),
            ListNumberingType.CharsCyrillicUpperLetterRu => BijectiveCapital(value, CyrillicUpperRu, CyrillicLowerRu),
            ListNumberingType.CharsCyrillicLowerLetterRu => Bijective(value, CyrillicLowerRu),
            ListNumberingType.CharsCyrillicUpperLetterNRu => RepeatedCapital(value, CyrillicUpperRu, CyrillicLowerRu),
            ListNumberingType.CharsCyrillicLowerLetterNRu => Repeated(value, CyrillicLowerRu),
            ListNumberingType.CharsCyrillicUpperLetterSr => BijectiveCapital(value, CyrillicUpperSr, CyrillicLowerSr),
            ListNumberingType.CharsCyrillicLowerLetterSr => Bijective(value, CyrillicLowerSr),
            ListNumberingType.CharsCyrillicUpperLetterNSr => RepeatedCapital(value, CyrillicUpperSr, CyrillicLowerSr),
            ListNumberingType.CharsCyrillicLowerLetterNSr => Repeated(value, CyrillicLowerSr),
            ListNumberingType.CharsCyrillicUpperLetterUk => BijectiveCapital(value, CyrillicUpperUk, CyrillicLowerUk),
            ListNumberingType.CharsCyrillicLowerLetterUk => Bijective(value, CyrillicLowerUk),
            ListNumberingType.CharsCyrillicUpperLetterNUk => RepeatedCapital(value, CyrillicUpperUk, CyrillicLowerUk),
            ListNumberingType.CharsCyrillicLowerLetterNUk => Repeated(value, CyrillicLowerUk),
            ListNumberingType.CharsGreekLowerLetter => Greek(value, GreekLower),
            ListNumberingType.CharsGreekUpperLetter => Greek(value, GreekUpper),
            ListNumberingType.CharsPersian => Bijective(value, PersianLetters),
            ListNumberingType.CharsPersianWord => PersianWord(value),
            ListNumberingType.SymbolChicago => Repeated(value, ChicagoSymbols),
            ListNumberingType.ArabicZero => digits.PadLeft(2, '0'),
            ListNumberingType.ArabicZero3 => digits.PadLeft(3, '0'),
            ListNumberingType.ArabicZero4 => digits.PadLeft(4, '0'),
            ListNumberingType.ArabicZero5 => digits.PadLeft(5, '0'),
            ListNumberingType.SzekelyRovas => OldHungarian(digits),
            _ => digits,
        };
    }

    // ── letter sequences (defaultnumberingprovider.cxx:400-437) ─────────────────────────────────────────────────────────

    /// <summary>lcl_formatChars: a..z, aa, ab, ... (bijective base n).</summary>
    private static string Bijective(int value, string table)
    {
        var result = new StringBuilder();
        for (int n = value - 1; ; n = (n - table.Length) / table.Length)
        {
            result.Insert(0, table[n % table.Length]);
            if (n < table.Length) break;
        }
        return result.ToString();
    }

    /// <summary>lcl_formatChars2: capital first letter, small ones after (A..Z, Aa, Ab).</summary>
    private static string BijectiveCapital(int value, string capitals, string smalls)
    {
        string small = Bijective(value, smalls);
        int first = smalls.IndexOf(small[0]);
        return capitals[first] + small.Substring(1);
    }

    /// <summary>lcl_formatChars1: a..z, aa, bb, ... (the letter repeated).</summary>
    private static string Repeated(int value, string table)
        => new(table[(value - 1) % table.Length], (value - 1) / table.Length + 1);

    /// <summary>lcl_formatChars3: A..Z, Aa, Bb, ...</summary>
    private static string RepeatedCapital(int value, string capitals, string smalls)
    {
        int index = (value - 1) % smalls.Length;
        return capitals[index] + new string(smalls[index], (value - 1) / smalls.Length);
    }

    /// <summary>The tableSize branch (:1059-1065): past the table the decimal number, unless the table recycles.</summary>
    private static string Table(int value, string table, bool recycle)
        => value > table.Length && !recycle ? value.ToString(CultureInfo.InvariantCulture) : table[(value - 1) % table.Length].ToString();

    private static string DigitByDigit(string digits, string table)
    {
        var result = new StringBuilder(digits.Length);
        foreach (char c in digits) result.Append(table[c - '0']);
        return result.ToString();
    }

    /// <summary>toRoman (:355): thousands as repeated M.</summary>
    public static string Roman(int value)
    {
        const string letters = "MDCLXVI--";
        var result = new StringBuilder();
        int mask = 1000, over = value / mask, position = 0;
        value -= over * mask;
        result.Append('M', over);
        while (mask > 0)
        {
            int number = value / mask, diff = 1;
            value %= mask;
            if (number > 5)
            {
                if (number < 9) result.Append(letters[position - 1]);
                diff++;
                number -= 5;
            }
            switch (number)
            {
                case 1: case 2: case 3: result.Append(letters[position], number); break;
                case 4: result.Append(letters[position]).Append(letters[position - diff]); break;
                case 5: result.Append(letters[position - diff]); break;
            }
            mask /= 10;
            position += 2;
        }
        return result.ToString();
    }

    // ── native numbers (nativenumbersupplier.cxx) ──────────────────────────────────────────────────────────────────────

    private sealed class NativeNumber(string digits, string multipliers, int flags, int[] exponents)
    {
        public readonly string Digits = digits, Multipliers = multipliers;
        public readonly int Flags = flags;
        public readonly int[] Exponents = exponents;
    }

    // nativenumbersupplier.cxx:27-41 flags
    private const int OmitZero = 1, OmitOnlyZero = 2, OmitOne = 0x1FC, OmitAll = OmitZero | OmitOne | OmitOnlyZero, OmitZeroOne = OmitZero | OmitOne;
    private static int OmitOneCheck(int bit) => 1 << (2 + bit);

    // numberchar.h:61-75 digits; :226-249 MultiplierExponent / MultiplierChar_6_CJK and _2_CJK
    private const string UpperZhDigits = "零壹贰叁肆伍陆柒捌玖";
    private const string UpperZhTwDigits = "零壹貳參肆伍陸柒捌玖";
    private const string TraditionalJaDigits = "零壱弐参四伍六七八九";
    private const string KoreanLowerDigits = "零一二三四五六七八九";
    private const string KoreanUpperDigits = "零壹貳參四伍六七八九";
    private const string HangulDigits = "영일이삼사오육칠팔구";
    private const string ArabicIndicDigits = "٠١٢٣٤٥٦٧٨٩";
    private const string EastArabicIndicDigits = "۰۱۲۳۴۵۶۷۸۹";
    private const string DevanagariDigits = "०१२३४५६७८९";
    private static readonly int[] Exponents6 = [12, 8, 4, 3, 2, 1], Exponents2 = [8, 4];

    // :412-421 natnum5 (zh_CN, zh_TW), :436-445 natnum8 (ja, ko), :447-448 natnum11
    private static readonly NativeNumber NatNum5SimplifiedChinese = new(UpperZhDigits, "兆亿万仟佰拾", 0, Exponents6);
    private static readonly NativeNumber NatNum5TraditionalChinese = new(UpperZhTwDigits, "兆億萬仟佰拾", 0, Exponents6);
    private static readonly NativeNumber NatNum8Japanese = new(TraditionalJaDigits, "億萬", OmitZeroOne, Exponents2);
    private static readonly NativeNumber NatNum8Korean = new(KoreanUpperDigits, "兆億萬阡佰拾", OmitAll, Exponents6);
    private static readonly NativeNumber NatNum11Korean = new(HangulDigits, "조억만천백십", OmitAll, Exponents6);

    /// <summary>AsciiToNative (:187) for a plain digit string.</summary>
    private static string AsciiToNative(string digits, NativeNumber number)
    {
        var output = new StringBuilder(digits.Length * 2);
        int length = digits.Length, last = number.Exponents.Length - 1;
        bool notZero = false;
        for (int begin = 0, end = length % number.Exponents[0]; end <= length; begin = end, end += number.Exponents[0])
        {
            if (end == 0) continue;
            int before = output.Length;
            notZero |= NumberMaker(digits, begin, end - begin, output, end == length ? -1 : 0, number);
            if (output.Length > 0 && number.Exponents[last] == 1 && output[output.Length - 1] == number.Digits[0]) output.Length--;
            if (notZero && before == output.Length && end != length) output.Append(number.Multipliers[0]);
        }
        if (!notZero && (number.Flags & OmitOnlyZero) == 0) output.Append(number.Digits[0]);
        return output.ToString();
    }

    /// <summary>AsciiToNative_numberMaker (:114).</summary>
    private static bool NumberMaker(string digits, int begin, int length, StringBuilder output, int multiplierIndex, NativeNumber number)
    {
        char multiplier = multiplierIndex == -1 ? '\0' : number.Multipliers[multiplierIndex];
        int last = number.Exponents.Length - 1;
        if (length <= number.Exponents[last])
        {
            if (number.Exponents[last] > 1)
            {
                bool any = false;
                for (int end = begin + length; begin < end; begin++)
                    if (any || digits[begin] != '0') { output.Append(number.Digits[digits[begin] - '0']); any = true; }
                if (any && multiplier != '\0') output.Append(multiplier);
                return any;
            }
            if (digits[begin] != '0')
            {
                if ((number.Flags & (multiplierIndex < 0 ? 0 : OmitOneCheck(multiplierIndex))) == 0 || digits[begin] != '1')
                    output.Append(number.Digits[digits[begin] - '0']);
                if (multiplier != '\0') output.Append(multiplier);
            }
            else if ((number.Flags & OmitZero) == 0 && output.Length > 0 && output[output.Length - 1] != number.Digits[0])
                output.Append(number.Digits[0]);
            return digits[begin] != '0';
        }
        bool printPower = false;
        for (int i = 1; i <= number.Exponents.Length; i++)
        {
            int part = length - (i == number.Exponents.Length ? 0 : number.Exponents[i]);
            if (part > 0)
            {
                printPower |= NumberMaker(digits, begin, part, output, i == number.Exponents.Length ? -1 : i, number);
                begin += part;
                length -= part;
            }
        }
        if (printPower)
        {
            if (output.Length > 0 && number.Exponents[last] == 1 && output[output.Length - 1] == number.Digits[0]) output.Length--;
            if (multiplier != '\0') output.Append(multiplier);
        }
        return printPower;
    }

    // :1026-1053 HebrewNumberCharArray (400 twice: the 15 / 16 substitution skips one entry)
    private static readonly (char letter, int value)[] HebrewNumerals =
    [
        ('ת', 400), ('ת', 400), ('ש', 300), ('ר', 200), ('ק', 100), ('צ', 90), ('פ', 80), ('ע', 70),
        ('ס', 60), ('נ', 50), ('מ', 40), ('ל', 30), ('כ', 20), ('י', 10), ('ט', 9), ('ח', 8),
        ('ז', 7), ('ו', 6), ('ה', 5), ('ד', 4), ('ג', 3), ('ב', 2), ('א', 1),
    ];

    /// <summary>NATNUM1 "he" (makeHebrewNumber :1055, no geresh): 1000 = the word thousand, 2000 = two + thousands.</summary>
    public static string HebrewNumber(int value)
    {
        var output = new StringBuilder();
        HebrewNumber(value, output, isLast: true);
        return output.ToString();
    }

    private static void HebrewNumber(long value, StringBuilder output, bool isLast)
    {
        int number = (int)(value % 1000);
        if (value > 1000)
        {
            HebrewNumber(value / 1000, output, number != 0);
            output.Append(' ');
        }
        if (number == 0)
        {
            output.Append(value == 1000 ? "אלף" : isLast ? "אלפים" : "אלפי");
            return;
        }
        for (int j = 0; number > 0 && j < HebrewNumerals.Length; j++)
        {
            if (number - HebrewNumerals[j].value < 0) continue;
            if (number is 15 or 16) j++;
            number -= HebrewNumerals[j].value;
            output.Append(HebrewNumerals[j].letter);
        }
    }

    // ── word numbers ───────────────────────────────────────────────────────────────────────────────────────────────────

    // :310-332 table_KoreanLegalWord_decade1 / decade2
    private static readonly string[] KoreanLegalOnes = ["하나", "둘", "셋", "넷", "다섯", "여섯", "일곱", "여덟", "아홉"];
    private static readonly string[] KoreanLegalTens = ["열", "스물", "서른", "마흔", "쉰", "예순", "일흔", "여든", "아흔"];

    /// <summary>lcl_formatKoreanLegalWord (:512), 1..99.</summary>
    public static string KoreanLegalWord(int value)
        => (value / 10 > 0 ? KoreanLegalTens[value / 10 - 1] : "") + (value % 10 > 0 ? KoreanLegalOnes[value % 10 - 1] : "");

    // :259-308 table_PersianWord_decade1 (0-19), decade2 (20-90), decade3 (100-900), decadeX (thousand, million, billion)
    private static readonly string[] PersianOnes =
    [
        "", "یک", "دو", "سه", "چهار", "پنج", "شش", "هفت", "هشت", "نه",
        "ده", "یازده", "دوازده", "سیزده", "چهارده",
        "پانزده", "شانزده", "هفده", "هجده", "نوزده",
    ];
    private static readonly string[] PersianTens =
        ["بیست", "سی", "چهل", "پنجاه", "شصت", "هفتاد", "هشتاد", "نود"];
    private static readonly string[] PersianHundreds =
        ["صد", "دویست", "سیصد", "چهارصد", "پانصد", "شصد", "هفتصد", "هشتصد", "نهصد"];
    private static readonly string[] PersianSections = ["هزار", "میلیون", "میلیارد"];
    private const string PersianAnd = " و ";

    /// <summary>lcl_formatPersianWord (:446): a zero group stops the walk, as LibreOffice's while loop does.</summary>
    private static string PersianWord(int value)
    {
        var text = new StringBuilder();
        int section = 0;
        for (int part; (part = value % 1000) != 0; value /= 1000, section++)
        {
            if (section > 0) text.Insert(0, PersianSections[section - 1] + PersianAnd);
            int digit = part % 100;
            if (digit < 20)
            {
                if (text.Length > 0) text.Insert(0, ' ');
                text.Insert(0, PersianOnes[digit]);
            }
            else
            {
                if ((digit = part % 10) != 0)
                {
                    if (text.Length > 0) text.Insert(0, PersianAnd);
                    text.Insert(0, PersianOnes[digit]);
                }
                if ((digit = part / 10 % 10) != 0)
                {
                    if (text.Length > 0) text.Insert(0, PersianAnd);
                    text.Insert(0, PersianTens[digit - 2]);
                }
            }
            if ((digit = part / 100) != 0)
            {
                if (text.Length > 0) text.Insert(0, PersianAnd);
                text.Insert(0, PersianHundreds[digit - 1]);
            }
        }
        return text.ToString();
    }

    // libnumbertext data/en.sor (the default rules: "one hundred one", "one thousand and one", "one hundred thousand, two hundred").
    // A Soros call strips the argument's leading zeros; "^" rules only match the whole input.
    private static readonly string[] EnglishOnes =
        ["", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen"];
    private static readonly string[] EnglishTens = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"];
    private static readonly string[] EnglishScales = ["thousand", "million", "billion"];

    private static string EnglishCardinal(string digits, bool top)
    {
        digits = digits.TrimStart('0');
        if (digits.Length == 0) return top ? "zero" : "";
        int length = digits.Length;
        if (length <= 2)
        {
            int n = int.Parse(digits, CultureInfo.InvariantCulture);
            if (n < 20) return EnglishOnes[n];
            return EnglishTens[n / 10] + (n % 10 != 0 ? "-" + EnglishOnes[n % 10] : "");
        }
        if (length == 3)
        {
            string rest = EnglishCardinal(digits.Substring(1), top: false);
            return EnglishOnes[digits[0] - '0'] + " hundred" + (rest.Length > 0 ? " " + rest : "");
        }
        if (length <= 5 && digits[length - 3] != '0')       // (\d{1,2})([1-9]\d\d) $1 thousand[ $2]
            return EnglishCardinal(digits.Substring(0, length - 3), top: false) + " thousand " + EnglishCardinal(digits.Substring(length - 3), top: false);
        int scale = (length - 4) / 3;                         // (\d{1,3})(\d{3k}) $1 <scale>$(:\2)$2
        int tail = 3 * (scale + 1);
        string low = digits.Substring(length - tail), lowText = EnglishCardinal(low, top: false);
        string separator = low.TrimStart('0').Length == 0 ? "" : low.TrimStart('0').Length <= 2 ? " and " : ", ";
        return EnglishCardinal(digits.Substring(0, length - tail), top: false) + " " + EnglishScales[scale] + separator + lowText;
    }

    /// <summary>en.sor "== ordinal ==": the cardinal's last word made ordinal.</summary>
    private static string EnglishOrdinal(string cardinal)
    {
        foreach ((string ending, string replacement) in new[] { ("one", "first"), ("two", "second"), ("three", "third"), ("five", "fifth"), ("eight", "eighth"), ("nine", "ninth"), ("twelve", "twelfth") })
            if (cardinal.EndsWith(ending, StringComparison.Ordinal)) return cardinal.Substring(0, cardinal.Length - ending.Length) + replacement;
        return cardinal.EndsWith("y", StringComparison.Ordinal) ? cardinal.Substring(0, cardinal.Length - 1) + "ieth" : cardinal + "th";
    }

    /// <summary>en.sor "== ordinal-number ==": 1st, 2nd, 3rd, 11th-19th, else th.</summary>
    private static string EnglishOrdinalNumber(int value)
    {
        int lastTwo = value % 100, last = value % 10;
        string suffix = lastTwo is >= 10 and <= 19 ? "th" : last switch { 1 => "st", 2 => "nd", 3 => "rd", _ => "th" };
        return value.ToString(CultureInfo.InvariantCulture) + suffix;
    }

    private static string Capitalize(string text) => text.Length == 0 ? text : char.ToUpperInvariant(text[0]) + text.Substring(1);

    // libnumbertext data/zh.sor (simplified, the default section)
    private const string ChineseDigits = "零一二三四五六七八九";
    private const char ChineseZero = '零', ChineseTen = '十', ChineseHundred = '百', ChineseThousand = '千';
    private const string ChineseGroups = "万亿兆";   // 10^4, 10^8, 10^12

    private static string ChineseText(string digits, bool top)
    {
        digits = digits.TrimStart('0');
        if (digits.Length == 0) return top ? ChineseZero.ToString() : "";
        int length = digits.Length;
        if (length == 1) return ChineseDigits[digits[0] - '0'].ToString();
        if (length == 2)
            return (top && digits[0] == '1' ? "" : ChineseDigits[digits[0] - '0'].ToString()) + ChineseTen + ChineseText(digits.Substring(1), top: false);
        if (length <= 4)
        {
            char unit = length == 3 ? ChineseHundred : ChineseThousand;
            string rest = digits.Substring(1), head = ChineseDigits[digits[0] - '0'].ToString() + unit;
            if (rest.TrimStart('0').Length == 0) return head;
            return head + (rest[0] == '0' ? ChineseZero.ToString() : "") + ChineseText(rest, top: false);
        }
        int group = Math.Min((length - 1) / 4, ChineseGroups.Length);      // (\d{1,4})0?(\d{4k})
        int tail = 4 * group;
        string low = digits.Substring(length - tail), high = ChineseText(digits.Substring(0, length - tail), top: false);
        if (low.TrimStart('0').Length == 0) return high + ChineseGroups[group - 1];
        return high + ChineseGroups[group - 1] + (low[0] == '0' ? ChineseZero.ToString() : "") + ChineseText(low, top: false);
    }

    // libnumbertext data/hu_Hung.sor: Old Hungarian numerals (one, five, ten, fifty, hundred, thousand)
    private const string RovasOne = "\U00010CFA", RovasFive = "\U00010CFB", RovasTen = "\U00010CFC", RovasFifty = "\U00010CFD",
                         RovasHundred = "\U00010CFE", RovasThousand = "\U00010CFF";

    private static string OldHungarian(string digits)
    {
        digits = digits.TrimStart('0');
        int length = digits.Length;
        if (length == 0) return "";
        if (length == 1) return OldHungarianDigit(digits[0] - '0', RovasOne, RovasFive);
        if (length == 2) return OldHungarianDigit(digits[0] - '0', RovasTen, RovasFifty) + OldHungarian(digits.Substring(1));
        if (length == 3) return (digits[0] == '1' ? "" : OldHungarian(digits.Substring(0, 1))) + RovasHundred + OldHungarian(digits.Substring(1));
        string thousands = digits.Substring(0, length - 3);
        return (thousands == "1" ? "" : OldHungarian(thousands)) + RovasThousand + OldHungarian(digits.Substring(length - 3));
    }

    private static string OldHungarianDigit(int digit, string one, string five)
    {
        var text = new StringBuilder();
        if (digit >= 5) { text.Append(five); digit -= 5; }
        for (int i = 0; i < digit; i++) text.Append(one);
        return text.ToString();
    }

    // ── Greek (:521-597) ────────────────────────────────────────────────────────────────────────────────────────────────

    private const string GreekUpper = "ΑΒΓΔΕϛΖΗΘΙΚΛΜΝΞΟΠϟΡΣΤΥΦΧΨΩϠ";
    private const string GreekLower = "αβγδεϛζηθικλμνξοπϟρστυφχψωϡ";

    private static string Greek(int value, string table)
    {
        var result = new StringBuilder();
        int myriadPower = 2;
        for (int divisor = 100000000; divisor > 1; divisor /= 10000, myriadPower--)
        {
            if (value <= divisor - 1) continue;
            result.Append('Μ', myriadPower);
            result.Append(GreekSmall(value / divisor, table));
            value %= divisor;
            if (value > 0) result.Append('.');
        }
        return result.Append(GreekSmall(value, table)).ToString();
    }

    private static string GreekSmall(int value, string table)
    {
        var result = new StringBuilder();
        for (int position = 0; value > 0; value /= 10, position++)
        {
            int digit = value % 10;
            if (digit == 0) continue;
            char sign = table[digit - 1 + 9 * (position % 3)];
            if (sign == 'ϛ') result.Insert(0, table.Substring(19, 2));
            else result.Insert(0, sign);
            if (position > 2) result.Insert(0, '͵');
        }
        return result.ToString();
    }

    // ── tables (defaultnumberingprovider.cxx:87-250, bullet.h) ──────────────────────────────────────────────────────────

    private const string UpperLetters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ", LowerLetters = "abcdefghijklmnopqrstuvwxyz";
    private const string ChicagoSymbols = "*†‡§";
    private const string ArabicLetters = "أبتثجحخدذرزسشصضطظعغفقكلمنهوى";
    private const string ArabicAbjadLetters = "ابجدهوزحطيكلمنسعفصقرشتثخذضظغ";
    private const string ThaiLetters = "กขคงจฉชซฌญฎฏฐฑฒณดตถทธนบปผฝพฟภมยรฤลฦวศษสหฬอฮ";
    private const string HebrewLetters = "אבגדהוזחטיכלמנסעפצקרשת";
    private const string NepaliLetters = "कखगघङचछजझञटठडढणतथदधनपफबभमयरलवशषसह";
    private const string KhmerLetters = "កខគឃងចឆជឈញដឋឌឍណតថទធនបផពភមយរលវសហឡអ";
    private const string LaoLetters = "ກຂຄງຈຊຍດຕຖທນບປຜຝພຟມຢຣລວສຫອຮຯຮໜໝ";
    private const string TibetanLetters = "ཀཁགངཅཆཇཉཏཐདནཔཕབམཙཚཛཝཞཟའཡརལཤསཧཨ";
    private const string MyanmarLetters = "ကခဂဃငစဆဇဈညဋဌဍဎဏတထဒဓနပဖဗဘမယရလဝသဟဠအ";
    private const string CyrillicUpperBg = "АБВГДЕЖЗИКЛМНОПРСТУФХЦЧШЩЪЮЯ";
    private const string CyrillicLowerBg = "абвгдежзиклмнопрстуфхцчшщъюя";
    private const string CyrillicUpperRu = "АБВГДЕЖЗИКЛМНОПРСТУФХЦЧШЩЫЭЮЯ";
    private const string CyrillicLowerRu = "абвгдежзиклмнопрстуфхцчшщыэюя";
    private const string CyrillicUpperSr = "АБВГДЂЕЖЗИЈКЛЉМНЊОПРСТЋУФХЦЧЏШ";
    private const string CyrillicLowerSr = "абвгдђежзијклљмнњопрстћуфхцчџш";
    private const string CyrillicUpperUk = "АБВГҐДЕЄЖЗИІЇЙКЛМНОПРСТУФХЦЧШЩЬЮЯ";
    private const string CyrillicLowerUk = "абвгґдеєжзиіїйклмнопрстуфхцчшщьюя";
    private const string PersianLetters = "آبپتثجچحخدذرزژسشصضطظعـفقکگلمنوهی";
    // bullet.h:226 table_CircledNumber (50), :349 table_TianGan_zh, :353 table_DiZi_zh, :314 / :332 circled Hangul
    private const string CircledNumbers = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"
        + "㉑㉒㉓㉔㉕㉖㉗㉘㉙㉚㉛㉜㉝㉞㉟㊱㊲㊳㊴㊵㊶㊷㊸㊹㊺㊻㊼㊽㊾㊿";
    private const string TianGan = "甲乙丙丁戊己庚辛壬癸";
    private const string DiZi = "子丑寅卯辰巳午未申酉戌亥";
    private const string HangulCircledJamo = "㉠㉡㉢㉣㉤㉥㉦㉧㉨㉩㉪㉫㉬㉭";
    private const string HangulCircledSyllables = "㉮㉯㉰㉱㉲㉳㉴㉵㉶㉷㉸㉹㉺㉻";
}

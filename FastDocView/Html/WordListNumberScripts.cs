using System.Globalization;
using System.Text;

namespace FastDocView.Html;

/// <summary>
/// Word's list number formats in other scripts: CJK / Korean counting and legal numerals, Thai / Vietnamese / Hindi counting
/// words, kana orders (aiueo, iroha), Korean letters (ganada, chosung), Korean digits and Word's custom formats (Greek
/// numerals, Turkish and Bulgarian letters, 001 / 0001 / 00001). Every rule matches Word 16's own ListFormat.ListString
/// (test-data/word-lists: list-formats-v1-*.expected.txt, and list-formats-words-v1.expected.txt = values 1..129 plus large
/// start values per format, make-list-formats-words.ps1). Character tables are LibreOffice's where it has them:
/// i18npool/inc/bullet.h:26 / :75 / :124 / :175 (kana), :279 / :296 (Hangul), i18npool/source/nativenumber/data/numberchar.h:61-75
/// (CJK digits) and :226-236 (multipliers), defaultnumberingprovider.cxx:310-332 (Korean legal words), :222-232 + :541 gr_smallNum
/// (Greek numerals), :157-170 (Bulgarian letters). Where LibreOffice differs, Word wins: LibreOffice draws all six counting
/// formats as NUMBER_LOWER_ZH (sw/source/writerfilter/dmapper/ConversionHelper.cxx:592-598), without Word's per-format rules for
/// "one" and zero; its Greek thousands sign is U+0375 (Word: a comma); its Bulgarian table has no short i (Word counts 29
/// letters); it reads only the three zero-padded custom formats (ConversionHelper.cxx:643 ConvertCustomNumberFormat). Thai,
/// Vietnamese and Hindi counting are not in LibreOffice (ConversionHelper.cxx:629-639, a TODO); the Hindi words are Word's
/// (they differ from CLDR's in places, 6 and 15 among them).
/// </summary>
public static class WordListNumberScripts
{
    /// <summary>The label for value in format (an ST_NumberFormat name or "custom:" + w:format), or null when the format is not one of these.</summary>
    public static string? Format(int value, string format) => format switch
    {
        "japaneseCounting" => Counting(value, CjkLowerDigits, '\0', CjkLowerUnits, JapaneseMyriads, CountingOne.Japanese),
        "japaneseLegal" => Counting(value, JapaneseTraditionalDigits, '\0', JapaneseTraditionalUnits, JapaneseTraditionalMyriads, CountingOne.Always),
        // Word 16: below 100 counted (10 = ten), from 100 digit by digit with a white circle for zero (100 = one circle circle)
        "taiwaneseCounting" or "chineseCounting" => value < 100
            ? Counting(value, CjkLowerDigits, '\0', CjkLowerUnits, ChineseSimplifiedMyriads, CountingOne.OnlyTeens)
            : DigitByDigit(value, (char)0x25CB + CjkLowerDigits.Substring(1)),
        "ideographLegalTraditional" => Counting(value, ChineseTraditionalUpperDigits, (char)0x96F6, ChineseUpperUnits, ChineseTraditionalMyriads, CountingOne.Always),
        "taiwaneseCountingThousand" => Counting(value, CjkLowerDigits, (char)0x96F6, CjkLowerUnits, ChineseTraditionalMyriads, CountingOne.OnlyTeens),
        "chineseLegalSimplified" => Counting(value, ChineseSimplifiedUpperDigits, (char)0x96F6, ChineseUpperUnits, ChineseLegalSimplifiedMyriads, CountingOne.Always),
        "chineseCountingThousand" => Counting(value, CjkLowerDigits, (char)0x3007, CjkLowerUnits, ChineseSimplifiedMyriads, CountingOne.OnlyTeens),
        "koreanCounting" => Counting(value, HangulDigits, '\0', HangulUnits, HangulMyriads, CountingOne.Korean),
        "koreanDigital" => value <= 0 ? DigitByDigit(value, HangulDigits) : ListNumberingTypes.Format(value, ListNumberingType.NumberDigitalKo),
        // 1..99 in words as LibreOffice's NUMBER_LEGAL_KO; from 100 Word counts with its own "one" rule, LibreOffice does not
        "koreanLegal" => value is >= 1 and <= 99 ? ListNumberingTypes.KoreanLegalWord(value) : Counting(value, HangulDigits, '\0', HangulUnits, HangulMyriads, CountingOne.Korean),
        "thaiCounting" => ThaiCounting(value),
        "vietnameseCounting" => VietnameseCounting(value),
        "hindiCounting" => HindiCounting(value),
        // the kana and Hangul orders cycle as LibreOffice's AIU / IROHA / HANGUL types (bRecycleSymbol)
        "aiueo" => Cycle(value, ListNumberingType.AiuHalfwidthJa),
        "aiueoFullWidth" => Cycle(value, ListNumberingType.AiuFullwidthJa),
        "iroha" => Cycle(value, ListNumberingType.IrohaHalfwidthJa),
        "irohaFullWidth" => Cycle(value, ListNumberingType.IrohaFullwidthJa),
        "ganada" => Cycle(value, ListNumberingType.HangulSyllableKo),
        "chosung" => Cycle(value, ListNumberingType.HangulJamoKo),
        CustomGreekLower => GreekNumeral(value, GreekLower),
        CustomGreekUpper => GreekNumeral(value, GreekUpper),
        CustomZero3 => value <= 0 ? ZeroPadded(value, 3) : ListNumberingTypes.Format(value, ListNumberingType.ArabicZero3),
        CustomZero4 => value <= 0 ? ZeroPadded(value, 4) : ListNumberingTypes.Format(value, ListNumberingType.ArabicZero4),
        CustomZero5 => value <= 0 ? ZeroPadded(value, 5) : ListNumberingTypes.Format(value, ListNumberingType.ArabicZero5),
        CustomTurkishLower => RepeatedLetter(value, TurkishLower),
        CustomTurkishUpper => RepeatedLetter(value, TurkishUpper),
        CustomBulgarianLower => RepeatedLetter(value, BulgarianLower),
        CustomBulgarianUpper => RepeatedLetter(value, BulgarianUpper),
        _ => null,
    };

    // Word's custom formats: "custom:" + the .docx w:numFmt val="custom" w:format Word 16 writes (MSONFC 60..68 in .doc / RTF)
    public const string CustomPrefix = "custom:";
    public const string CustomGreekLower = CustomPrefix + "\u03B1, \u03B2, \u03B3, ...";
    public const string CustomGreekUpper = CustomPrefix + "\u0391, \u0392, \u0393, ...";
    public const string CustomZero3 = CustomPrefix + "001, 002, 003, ...";
    public const string CustomZero4 = CustomPrefix + "0001, 0002, 0003, ...";
    public const string CustomZero5 = CustomPrefix + "00001, 00002, 00003, ...";
    public const string CustomTurkishLower = CustomPrefix + "a, \u00E7, \u011D, ...";
    public const string CustomTurkishUpper = CustomPrefix + "A, \u00C7, \u011C, ...";
    public const string CustomBulgarianLower = CustomPrefix + "\u0430, \u0439, \u043A, ...";
    public const string CustomBulgarianUpper = CustomPrefix + "\u0410, \u0419, \u041A, ...";

    /// <summary>MSONFC 60..68 in order (Word's WdListNumberStyle: Greek lower / upper, 001, 0001, 00001, Turkish lower / upper, Bulgarian lower / upper).</summary>
    public static readonly string[] CustomFormatsByCode =
        [CustomGreekLower, CustomGreekUpper, CustomZero3, CustomZero4, CustomZero5, CustomTurkishLower, CustomTurkishUpper, CustomBulgarianLower, CustomBulgarianUpper];

    // ── CJK / Korean counting ─────────────────────────────────────────────────────────────────────────────────────────────

    /// <summary>Where Word leaves out "one" before a multiplier (Word 16 ListString, list-formats-words-v1).</summary>
    private enum CountingOne
    {
        /// <summary>Never: 10 = one ten (legal formats).</summary>
        Always,
        /// <summary>Only for 10..19 (Chinese / Taiwanese counting): 10 = ten, 100000 = one ten myriad.</summary>
        OnlyTeens,
        /// <summary>Before ten and hundred always, before thousand only in the leading four digits, never before a myriad (Japanese).</summary>
        Japanese,
        /// <summary>As Japanese, and before a myriad whose multiplier is one (Korean: 10000 = man).</summary>
        Korean,
    }

    // numberchar.h:64 Chinese lower (index 0 = U+3007), :65 S. Chinese upper, :66 T. Chinese upper, :67 Japanese traditional, :70 Korean Hangul
    private const string CjkLowerDigits = "\u3007\u4E00\u4E8C\u4E09\u56DB\u4E94\u516D\u4E03\u516B\u4E5D";
    private const string ChineseSimplifiedUpperDigits = "\u96F6\u58F9\u8D30\u53C1\u8086\u4F0D\u9646\u67D2\u634C\u7396";
    private const string ChineseTraditionalUpperDigits = "\u96F6\u58F9\u8CB3\u53C3\u8086\u4F0D\u9678\u67D2\u634C\u7396";
    private const string JapaneseTraditionalDigits = "\u96F6\u58F1\u5F10\u53C2\u56DB\u4F0D\u516D\u4E03\u516B\u4E5D";
    private const string HangulDigits = "\uC601\uC77C\uC774\uC0BC\uC0AC\uC624\uC721\uCE60\uD314\uAD6C";
    // numberchar.h:226-236 MultiplierChar_6_CJK: ten, hundred, thousand; then myriad, hundred million
    private const string CjkLowerUnits = "\u5341\u767E\u5343";
    private const string ChineseUpperUnits = "\u62FE\u4F70\u4EDF";
    private const string JapaneseTraditionalUnits = "\u62FE\u767E\u9621";
    private const string HangulUnits = "\uC2ED\uBC31\uCC9C";
    private const string ChineseSimplifiedMyriads = "\u4E07\u4EBF";
    private const string ChineseTraditionalMyriads = "\u842C\u5104";
    private const string ChineseLegalSimplifiedMyriads = "\u842C\u4EBF";   // Word 16: 10000 = one + U+842C (LibreOffice: U+4E07)
    private const string JapaneseMyriads = "\u4E07\u5104";
    private const string JapaneseTraditionalMyriads = "\u842C\u5104";
    private const string HangulMyriads = "\uB9CC\uC5B5";

    /// <summary>
    /// Multiplier numerals: each non-zero digit, then its multiplier within a group of four (ten, hundred, thousand), then the
    /// group's myriad. With a zero sign, a run of zeros between non-zero digits is written once (Word: 101 = one hundred zero
    /// one, 10010 = one myriad zero one ten); without one, zeros are dropped.
    /// </summary>
    private static string Counting(int value, string digits, char zero, string units, string myriads, CountingOne one)
    {
        if (value <= 0) return value.ToString(CultureInfo.InvariantCulture);
        string plain = value.ToString(CultureInfo.InvariantCulture);
        int length = plain.Length, leadingGroup = (length - 1) / 4;
        var result = new StringBuilder(length * 2);
        bool pendingZero = false;
        for (int i = 0; i < length; i++)
        {
            int digit = plain[i] - '0', position = length - 1 - i, unit = position % 4, group = position / 4;
            int groupValue = group == 0 ? 0 : value / Power10(4 * group) % 10000;
            if (digit == 0)
            {
                if (result.Length > 0) pendingZero = true;
            }
            else
            {
                if (pendingZero && zero != '\0') result.Append(zero);
                pendingZero = false;
                bool omitOne = digit == 1 && (unit > 0
                    ? one switch
                    {
                        CountingOne.OnlyTeens => value < 20,
                        CountingOne.Japanese or CountingOne.Korean => unit < 3 || group == leadingGroup,
                        _ => false,
                    }
                    : one == CountingOne.Korean && group > 0 && groupValue == 1);
                if (!omitOne) result.Append(digits[digit]);
                if (unit > 0) result.Append(units[unit - 1]);
            }
            if (unit == 0 && group > 0 && groupValue != 0) result.Append(myriads[group - 1]);
        }
        return result.ToString();
    }

    private static int Power10(int exponent)
    {
        int power = 1;
        for (int i = 0; i < exponent; i++) power *= 10;
        return power;
    }

    private static string DigitByDigit(int value, string digits)
    {
        string plain = value.ToString(CultureInfo.InvariantCulture);
        var result = new StringBuilder(plain.Length);
        foreach (char c in plain) result.Append(c is >= '0' and <= '9' ? digits[c - '0'] : c);
        return result.ToString();
    }

    // ── Thai, Vietnamese, Hindi counting words ────────────────────────────────────────────────────────────────────────────

    private static readonly string[] ThaiDigitWords =
        ["", "\u0E2B\u0E19\u0E36\u0E48\u0E07", "\u0E2A\u0E2D\u0E07", "\u0E2A\u0E32\u0E21", "\u0E2A\u0E35\u0E48", "\u0E2B\u0E49\u0E32",
         "\u0E2B\u0E01", "\u0E40\u0E08\u0E47\u0E14", "\u0E41\u0E1B\u0E14", "\u0E40\u0E01\u0E49\u0E32"];
    // -, ten, hundred, thousand, ten thousand, hundred thousand; then million
    private static readonly string[] ThaiUnits =
        ["", "\u0E2A\u0E34\u0E1A", "\u0E23\u0E49\u0E2D\u0E22", "\u0E1E\u0E31\u0E19", "\u0E2B\u0E21\u0E37\u0E48\u0E19", "\u0E41\u0E2A\u0E19"];
    private const string ThaiMillion = "\u0E25\u0E49\u0E32\u0E19";
    private const string ThaiOneAfterHigher = "\u0E40\u0E2D\u0E47\u0E14";   // et: a final one after higher digits (11, 101, 1001)
    private const string ThaiTwoInTens = "\u0E22\u0E35\u0E48";              // yi: 20..29

    /// <summary>Word's Thai counting: 10 = sip (no one), 2x = yi sip, a final 1 after any higher digit = et; millions recurse.</summary>
    private static string ThaiCounting(int value)
    {
        if (value <= 0) return value.ToString(CultureInfo.InvariantCulture);
        if (value >= 1000000)
            return ThaiCounting(value / 1000000) + ThaiMillion + ThaiBelowMillion(value % 1000000, hasHigher: true);
        return ThaiBelowMillion(value, hasHigher: false);
    }

    private static string ThaiBelowMillion(int value, bool hasHigher)
    {
        if (value == 0) return "";
        string plain = value.ToString(CultureInfo.InvariantCulture);
        var result = new StringBuilder();
        for (int i = 0; i < plain.Length; i++)
        {
            int digit = plain[i] - '0', position = plain.Length - 1 - i;
            if (digit == 0) continue;
            if (position == 1) result.Append(digit == 1 ? "" : digit == 2 ? ThaiTwoInTens : ThaiDigitWords[digit]).Append(ThaiUnits[1]);
            else if (position == 0 && digit == 1 && (hasHigher || value > 10)) result.Append(ThaiOneAfterHigher);
            else result.Append(ThaiDigitWords[digit]).Append(ThaiUnits[position]);
        }
        return result.ToString();
    }

    // Word writes the tone mark as a combining character after the vowel (one = m, o-circumflex, U+0323, t)
    private static readonly string[] VietnameseDigitWords =
        ["", "m\u00F4\u0323t", "hai", "ba", "b\u00F4\u0301n", "n\u0103m", "sa\u0301u", "ba\u0309y", "ta\u0301m", "chi\u0301n"];
    private const string VietnameseTen = "m\u01B0\u01A1\u0300i";        // 10..19
    private const string VietnameseTens = "m\u01B0\u01A1i";             // after 2..9
    private const string VietnameseOneAfterTens = "m\u00F4\u0301t";    // 21, 31...
    private const string VietnameseFiveAfterTens = "l\u0103m";         // 15, 25...
    private const string VietnameseHundred = "tr\u0103m";
    private const string VietnameseZeroTens = "le\u0309";              // 101 = one hundred le one
    private const string VietnameseThousand = "nga\u0300n";

    /// <summary>Word's Vietnamese counting (Word accepts start values up to 1000; past that the thousands recurse).</summary>
    private static string VietnameseCounting(int value)
    {
        if (value <= 0) return value.ToString(CultureInfo.InvariantCulture);
        if (value >= 1000)
        {
            int rest = value % 1000;
            return VietnameseCounting(value / 1000) + " " + VietnameseThousand + (rest == 0 ? "" : " " + VietnameseBelowThousand(rest));
        }
        return VietnameseBelowThousand(value);
    }

    private static string VietnameseBelowThousand(int value)
    {
        var words = new List<string>(5);
        int hundreds = value / 100, tens = value / 10 % 10, ones = value % 10;
        if (hundreds > 0)
        {
            words.Add(VietnameseDigitWords[hundreds]);
            words.Add(VietnameseHundred);
            if (tens == 0 && ones > 0) words.Add(VietnameseZeroTens);
        }
        if (tens == 1) words.Add(VietnameseTen);
        else if (tens > 1) { words.Add(VietnameseDigitWords[tens]); words.Add(VietnameseTens); }
        if (ones > 0)
            words.Add(ones == 5 && tens > 0 ? VietnameseFiveAfterTens : ones == 1 && tens > 1 ? VietnameseOneAfterTens : VietnameseDigitWords[ones]);
        return string.Join(" ", words);
    }

    // Word 16's ListString for hindiCounting 1..99 (list-formats-words-v1.expected.txt)
    private static readonly string[] HindiWords =
    [
        "\u090F\u0915", "\u0926\u094B", "\u0924\u0940\u0928", "\u091A\u093E\u0930", "\u092A\u093E\u0901\u091A", "\u091B\u0903", "\u0938\u093E\u0924", "\u0906\u0920", "\u0928\u094C",
        "\u0926\u0938", "\u0917\u094D\u092F\u093E\u0930\u0939", "\u092C\u093E\u0930\u0939", "\u0924\u0947\u0930\u0939", "\u091A\u094C\u0926\u0939", "\u092A\u0902\u0926\u094D\u0930\u0939", "\u0938\u094B\u0932\u0939", "\u0938\u0924\u094D\u0930\u0939", "\u0905\u0920\u094D\u0920\u093E\u0930\u0939", "\u0909\u0928\u094D\u0928\u0940\u0938",
        "\u092C\u0940\u0938", "\u0907\u0915\u094D\u0915\u0940\u0938", "\u092C\u093E\u0908\u0938", "\u0924\u0947\u0908\u0938", "\u091A\u094C\u092C\u0940\u0938", "\u092A\u091A\u094D\u091A\u0940\u0938", "\u091B\u092C\u094D\u092C\u0940\u0938", "\u0938\u0924\u094D\u0924\u093E\u0908\u0938", "\u0905\u0920\u094D\u0920\u093E\u0908\u0938", "\u0909\u0928\u094D\u0924\u0940\u0938",
        "\u0924\u0940\u0938", "\u0907\u0915\u094D\u0924\u0940\u0938", "\u092C\u0924\u094D\u0924\u0940\u0938", "\u0924\u0948\u0902\u0924\u0940\u0938", "\u091A\u094C\u0902\u0924\u0940\u0938", "\u092A\u0948\u0902\u0924\u0940\u0938", "\u091B\u0924\u0940\u0938", "\u0938\u0948\u0902\u0924\u0940\u0938", "\u0905\u0921\u093C\u0924\u0940\u0938", "\u0909\u0928\u094D\u0924\u093E\u0932\u0940\u0938",
        "\u091A\u093E\u0932\u0940\u0938", "\u0907\u0915\u094D\u0924\u093E\u0932\u0940\u0938", "\u092C\u092F\u093E\u0932\u0940\u0938", "\u0924\u0948\u0924\u093E\u0932\u0940\u0938", "\u091A\u0935\u093E\u0932\u0940\u0938", "\u092A\u0948\u0902\u0924\u093E\u0932\u0940\u0938", "\u091B\u093F\u092F\u093E\u0932\u0940\u0938", "\u0938\u0924\u094D\u0924\u093E\u0932\u0940\u0938", "\u0905\u095C\u0924\u093E\u0932\u0940\u0938", "\u0909\u0928\u094D\u091A\u093E\u0938",
        "\u092A\u091A\u093E\u0938", "\u0907\u0915\u093E\u0935\u0928", "\u092C\u093E\u0935\u0928", "\u0924\u093F\u0930\u092A\u0928", "\u091A\u094C\u0935\u0928", "\u092A\u091A\u092A\u0928", "\u091B\u092A\u094D\u092A\u0928", "\u0938\u0924\u094D\u0924\u093E\u0935\u0928", "\u0905\u0920\u094D\u0920\u093E\u0935\u0928", "\u0909\u0928\u0938\u0920",
        "\u0938\u093E\u0920", "\u0907\u0915\u0938\u0920", "\u092C\u093E\u0938\u0920", "\u0924\u093F\u0930\u0938\u0920", "\u091A\u094C\u0902\u0938\u0920", "\u092A\u0948\u0902\u0938\u0920", "\u091B\u093F\u092F\u093E\u0938\u0920", "\u0938\u095C\u0938\u0920", "\u0905\u095C\u0938\u0920", "\u0909\u0928\u0939\u0924\u094D\u0924\u0930",
        "\u0938\u0924\u094D\u0924\u0930", "\u090F\u0915\u0939\u0924\u094D\u0924\u0930", "\u092C\u0939\u0924\u094D\u0924\u0930", "\u0924\u093F\u0939\u0924\u094D\u0924\u0930", "\u091A\u094C\u0939\u0924\u094D\u0924\u0930", "\u092A\u091B\u0924\u094D\u0924\u0930", "\u091B\u093F\u0939\u0924\u094D\u0924\u0930", "\u0938\u0924\u0939\u0924\u094D\u0924\u0930", "\u0905\u0920\u0939\u0924\u094D\u0924\u0930", "\u0909\u0928\u094D\u0928\u093E\u0938\u0940",
        "\u0905\u0938\u094D\u0938\u0940", "\u0907\u0915\u094D\u092F\u093E\u0938\u0940", "\u092C\u092F\u093E\u0938\u0940", "\u0924\u093F\u0930\u093E\u0938\u0940", "\u091A\u094C\u0930\u093E\u0938\u0940", "\u092A\u091A\u093E\u0938\u0940", "\u091B\u093F\u092F\u093E\u0938\u0940", "\u0938\u0924\u094D\u0924\u093E\u0938\u0940", "\u0905\u0920\u094D\u0920\u093E\u0938\u0940", "\u0928\u0935\u093E\u0938\u0940",
        "\u0928\u092C\u094D\u092C\u0947", "\u0907\u0915\u094D\u092F\u093E\u0928\u0935\u0947", "\u092C\u093E\u0928\u0935\u0947", "\u0924\u093F\u0930\u093E\u0928\u0935\u0947", "\u091A\u094C\u0930\u093E\u0928\u0935\u0947", "\u092A\u091B\u093E\u0928\u0935\u0947", "\u091B\u094D\u092F\u093E\u0928\u0935\u0947", "\u0938\u0924\u094D\u0924\u093E\u0928\u0935\u0947", "\u0905\u0920\u094D\u0920\u093E\u0928\u0935\u0947", "\u0928\u093F\u0928\u094D\u092F\u093E\u0928\u0935\u0947"
    ];
    private const string HindiHundred = "\u0938\u094C";
    private const string HindiThousand = "\u0939\u091C\u093E\u0930";

    /// <summary>Word's Hindi counting: the word for 1..99, "N sau", "N hajar", separated by spaces (Word accepts start values below 10000).</summary>
    private static string HindiCounting(int value)
    {
        if (value <= 0) return value.ToString(CultureInfo.InvariantCulture);
        var words = new List<string>(5);
        if (value >= 1000) { words.Add(HindiCounting(value / 1000)); words.Add(HindiThousand); value %= 1000; }
        if (value >= 100) { words.Add(HindiWords[value / 100 - 1]); words.Add(HindiHundred); value %= 100; }
        if (value > 0) words.Add(HindiWords[value - 1]);
        return string.Join(" ", words);
    }

    // ── Letter orders ─────────────────────────────────────────────────────────────────────────────────────────────────────

    // bullet.h:75 table_AIUHalfWidth_ja_JP, :26 table_AIUFullWidth_ja_JP, :175 table_IROHAHalfWidth_ja_JP (wi / we have no
    // half-width form: full-width ones), :124 table_IROHAFullWidth_ja_JP, :296 table_HangulSyllable_ko, :279 table_HangulJamo_ko.
    // LibreOffice cycles them (bRecycleSymbol, defaultnumberingprovider.cxx:829-858, :1054-1057) and so does Word.
    internal const string AiueoHalfWidth =
        "\uFF71\uFF72\uFF73\uFF74\uFF75\uFF76\uFF77\uFF78\uFF79\uFF7A\uFF7B\uFF7C\uFF7D\uFF7E\uFF7F\uFF80\uFF81\uFF82\uFF83\uFF84\uFF85\uFF86\uFF87"
        + "\uFF88\uFF89\uFF8A\uFF8B\uFF8C\uFF8D\uFF8E\uFF8F\uFF90\uFF91\uFF92\uFF93\uFF94\uFF95\uFF96\uFF97\uFF98\uFF99\uFF9A\uFF9B\uFF9C\uFF66\uFF9D";
    internal const string AiueoFullWidth =
        "\u30A2\u30A4\u30A6\u30A8\u30AA\u30AB\u30AD\u30AF\u30B1\u30B3\u30B5\u30B7\u30B9\u30BB\u30BD\u30BF\u30C1\u30C4\u30C6\u30C8\u30CA\u30CB\u30CC"
        + "\u30CD\u30CE\u30CF\u30D2\u30D5\u30D8\u30DB\u30DE\u30DF\u30E0\u30E1\u30E2\u30E4\u30E6\u30E8\u30E9\u30EA\u30EB\u30EC\u30ED\u30EF\u30F2\u30F3";
    internal const string IrohaHalfWidth =
        "\uFF72\uFF9B\uFF8A\uFF86\uFF8E\uFF8D\uFF84\uFF81\uFF98\uFF87\uFF99\uFF66\uFF9C\uFF76\uFF96\uFF80\uFF9A\uFF7F\uFF82\uFF88\uFF85\uFF97\uFF91\uFF73"
        + "\u30F0\uFF89\uFF75\uFF78\uFF94\uFF8F\uFF79\uFF8C\uFF7A\uFF74\uFF83\uFF71\uFF7B\uFF77\uFF95\uFF92\uFF90\uFF7C\u30F1\uFF8B\uFF93\uFF7E\uFF7D\uFF9D";
    internal const string IrohaFullWidth =
        "\u30A4\u30ED\u30CF\u30CB\u30DB\u30D8\u30C8\u30C1\u30EA\u30CC\u30EB\u30F2\u30EF\u30AB\u30E8\u30BF\u30EC\u30BD\u30C4\u30CD\u30CA\u30E9\u30E0\u30A6"
        + "\u30F0\u30CE\u30AA\u30AF\u30E4\u30DE\u30B1\u30D5\u30B3\u30A8\u30C6\u30A2\u30B5\u30AD\u30E6\u30E1\u30DF\u30B7\u30F1\u30D2\u30E2\u30BB\u30B9\u30F3";
    internal const string HangulSyllables = "\uAC00\uB098\uB2E4\uB77C\uB9C8\uBC14\uC0AC\uC544\uC790\uCC28\uCE74\uD0C0\uD30C\uD558";
    internal const string HangulJamo = "\u3131\u3134\u3137\u3139\u3141\u3142\u3145\u3147\u3148\u314A\u314B\u314C\u314D\u314E";

    private static string Cycle(int value, ListNumberingType type) => value <= 0 ? value.ToString(CultureInfo.InvariantCulture) : ListNumberingTypes.Format(value, type);

    // Word's Turkish letters (29, dotless and dotted i); Word's Bulgarian letters: defaultnumberingprovider.cxx:157-170 plus the short i Word counts (29)
    private const string TurkishLower = "abc\u00E7defg\u011Fh\u0131ijklmno\u00F6prs\u015Ftu\u00FCvyz";
    private const string TurkishUpper = "ABC\u00C7DEFG\u011EHI\u0130JKLMNO\u00D6PRS\u015ETU\u00DCVYZ";
    private const string BulgarianLower = "\u0430\u0431\u0432\u0433\u0434\u0435\u0436\u0437\u0438\u0439\u043A\u043B\u043C\u043D\u043E\u043F\u0440\u0441\u0442\u0443\u0444\u0445\u0446\u0447\u0448\u0449\u044A\u044E\u044F";
    private const string BulgarianUpper = "\u0410\u0411\u0412\u0413\u0414\u0415\u0416\u0417\u0418\u0419\u041A\u041B\u041C\u041D\u041E\u041F\u0420\u0421\u0422\u0423\u0424\u0425\u0426\u0427\u0428\u0429\u042A\u042E\u042F";

    /// <summary>Word's letter lists: the letter of (n - 1) mod count, written (n - 1) / count + 1 times.</summary>
    private static string RepeatedLetter(int value, string letters)
        => value <= 0 ? "" : new string(letters[(value - 1) % letters.Length], (value - 1) / letters.Length + 1);

    private static string ZeroPadded(int value, int width) => value.ToString(new string('0', width), CultureInfo.InvariantCulture);

    // defaultnumberingprovider.cxx:222-232: units (stigma at 6), tens (koppa at 90), hundreds (sampi at 900); Word 16 writes the
    // capital koppa U+03DE in upper case (LibreOffice: the small U+03DF in both)
    private const string GreekUpper = "\u0391\u0392\u0393\u0394\u0395\u03DB\u0396\u0397\u0398\u0399\u039A\u039B\u039C\u039D\u039E\u039F\u03A0\u03DE\u03A1\u03A3\u03A4\u03A5\u03A6\u03A7\u03A8\u03A9\u03E0";
    private const string GreekLower = "\u03B1\u03B2\u03B3\u03B4\u03B5\u03DB\u03B6\u03B7\u03B8\u03B9\u03BA\u03BB\u03BC\u03BD\u03BE\u03BF\u03C0\u03DF\u03C1\u03C3\u03C4\u03C5\u03C6\u03C7\u03C8\u03C9\u03E1";
    private const char GreekMyriad = (char)0x039C;   // defaultnumberingprovider.cxx:528 MYRIAD_SYM

    /// <summary>
    /// Greek numerals as LibreOffice's gr_smallNum (defaultnumberingprovider.cxx:541): a letter per non-zero digit, stigma
    /// written sigma + tau (Word too: 6 = sigma tau), a thousands digit marked before its letter - Word with a comma (1234 =
    /// comma alpha sigma lambda delta), LibreOffice with U+0375. Past 9999 LibreOffice's myriad form (:569 lcl_formatCharsGR);
    /// Word refuses such start values.
    /// </summary>
    private static string GreekNumeral(int value, string table)
    {
        if (value <= 0) return value.ToString(CultureInfo.InvariantCulture);
        var result = new StringBuilder();
        int myriadPower = 2;
        for (int divisor = 100000000; divisor > 1; divisor /= 10000, myriadPower--)
        {
            if (value < divisor) continue;
            result.Append(GreekMyriad, myriadPower);
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
            int index = digit - 1 + 9 * (position % 3);
            string sign = table[index] == (char)0x03DB ? table.Substring(19, 2) : table[index].ToString();   // :553-556 stigma -> sigma tau
            result.Insert(0, position > 2 ? "," + sign : sign);
        }
        return result.ToString();
    }
}

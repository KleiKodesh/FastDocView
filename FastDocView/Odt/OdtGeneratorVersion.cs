namespace FastDocView.Odt;

/// <summary>
/// Which product and version wrote an OpenDocument file, from its meta:generator, classified exactly as LibreOffice's import
/// does (LibreOffice/core master read 2026-11-04):
/// - xmloff/source/meta/xmlmetai.cxx:232-298 SvXMLMetaDocumentContext::setBuildId - the generator becomes a build id
///   "UPD$build" (the text between the first '/' after the first space and the next 'm', then what follows "$Build-"),
///   StarOffice / StarSuite 6-7 and OpenOffice.org 1 become "645$8687", NeoOffice/2 "680$9134", and a generator holding
///   "LibreOffice_project/" gets ";" + the digits and dots after its first '/';
/// - xmloff/source/core/xmlimp.cxx:171-338 DocumentInfo - a LibreOffice version after ';' decides alone (3.x ... 7.x, then
///   year.month), else the UPD / build numbers (getBuildIds :2075-2093) name the OpenOffice.org / AOO release;
/// - include/xmloff/xmlimp.hxx:529-564 the codes; xmlimp.cxx:2102-2112 isGeneratorVersionOlderThan.
/// </summary>
internal static class OdtGeneratorVersion
{
    public const ushort OOo1 = 10, OOo2 = 20, OOo30 = 30, OOo31 = 31, OOo32 = 32, OOo33 = 33, OOo34 = 34, Aoo40 = 40, Aoo4 = 41;
    public const ushort LibreOfficeFlag = 0x100;
    public const ushort LibreOffice3 = 30 | LibreOfficeFlag, LibreOffice41 = 41 | LibreOfficeFlag, LibreOffice42 = 42 | LibreOfficeFlag,
        LibreOffice43 = 43 | LibreOfficeFlag, LibreOffice44 = 44 | LibreOfficeFlag, LibreOffice5 = 50 | LibreOfficeFlag, LibreOffice6 = 60 | LibreOfficeFlag,
        LibreOffice63 = 63 | LibreOfficeFlag, LibreOffice7 = 70 | LibreOfficeFlag, LibreOffice76 = 76 | LibreOfficeFlag, LibreOffice242 = 80 | LibreOfficeFlag,
        LibreOffice248 = 81 | LibreOfficeFlag, LibreOffice252 = 82 | LibreOfficeFlag, LibreOffice258 = 83 | LibreOfficeFlag, LibreOffice262 = 84 | LibreOfficeFlag,
        LibreOffice268 = 85 | LibreOfficeFlag, LibreOfficeNew = 100 | LibreOfficeFlag;
    public const ushort Unknown = ushort.MaxValue;

    /// <summary>The generator's version code (<see cref="Unknown"/> without a generator or for other products).</summary>
    public static ushort Of(string? generator)
    {
        string buildIds = BuildIds(generator ?? "");
        if (buildIds.Length == 0) return Unknown;
        int semicolon = buildIds.IndexOf(';');
        if (semicolon >= 0 && semicolon + 1 < buildIds.Length) return LibreOfficeVersion(buildIds.Substring(semicolon + 1));
        int dollar = buildIds.IndexOf('$');
        if (dollar < 0) return Unknown;
        int upd = LeadingInt(buildIds.Substring(0, dollar));
        int end = buildIds.IndexOf(';', dollar);
        int build = LeadingInt(end < 0 ? buildIds.Substring(dollar + 1) : buildIds.Substring(dollar + 1, end - dollar - 1));
        if (upd >= 640 && upd <= 645) return OOo1;
        if (upd == 680) return OOo2;
        if (upd == 300 && build <= 9379) return OOo30;
        if (upd == 310) return OOo31;
        if (upd == 320) return OOo32;
        if (upd == 330) return OOo33;
        if (upd == 340) return OOo34;
        if (upd == 400 || upd == 401) return Aoo40;
        if (upd >= 410) return Aoo4;
        return Unknown;
    }

    /// <summary>isGeneratorVersionOlderThan (xmlimp.cxx:2102-2112): LibreOffice codes compare with the LibreOffice bound, the rest with the OOo bound.</summary>
    public static bool IsOlderThan(ushort version, ushort openOfficeVersion, ushort libreOfficeVersion)
        => (version & LibreOfficeFlag) != 0 ? version < libreOfficeVersion : version < openOfficeVersion;

    /// <summary>
    /// tdf#89475: a draw:angle / draw:rotation of a gradient without a unit is tenths of a degree in ODF older than 1.2, and in
    /// ODF 1.2 written by OpenOffice.org / AOO (any) or LibreOffice before 7; degrees otherwise (xmloff GradientStyle.cxx:115-128,
    /// TransGradientStyle.cxx:128-141; an unknown generator counts as new - its code carries the LibreOffice flag).
    /// </summary>
    public static bool GradientAngleIsTenths(double odfVersion, ushort generator)
        => odfVersion < 1.2 || (odfVersion == 1.2 && (IsOlderThan(generator, Aoo4, LibreOffice7) || generator == Aoo4));

    private static string BuildIds(string generator)
    {
        string buildId = "";
        int begin = generator.IndexOf(' ');
        if (begin >= 0)
        {
            begin = generator.IndexOf('/', begin);
            if (begin >= 0)
            {
                int end = generator.IndexOf('m', begin);
                if (end >= 0)
                {
                    const string buildCompare = "$Build-";
                    int build = generator.IndexOf(buildCompare, end, StringComparison.Ordinal);
                    if (build >= 0) buildId = generator.Substring(begin + 1, end - begin - 1) + "$" + generator.Substring(build + buildCompare.Length);
                }
            }
        }
        if (buildId.Length == 0)
        {
            if (generator.StartsWith("StarOffice 7", StringComparison.Ordinal) || generator.StartsWith("StarSuite 7", StringComparison.Ordinal)
                || generator.StartsWith("StarOffice 6", StringComparison.Ordinal) || generator.StartsWith("StarSuite 6", StringComparison.Ordinal)
                || generator.StartsWith("OpenOffice.org 1", StringComparison.Ordinal))
                buildId = "645$8687";
            else if (generator.StartsWith("NeoOffice/2", StringComparison.Ordinal))
                buildId = "680$9134";
        }
        if (generator.IndexOf("LibreOffice_project/", StringComparison.Ordinal) >= 0)
        {
            int firstSlash = generator.IndexOf('/');
            int i = firstSlash + 1;
            while (i < generator.Length && (generator[i] is >= '0' and <= '9' || generator[i] == '.')) i++;
            if (i > firstSlash + 1) buildId += ";" + generator.Substring(firstSlash + 1, i - firstSlash - 1);
        }
        return buildId;
    }

    /// <summary>DocumentInfo's LibreOffice branch (xmlimp.cxx:183-288).</summary>
    private static ushort LibreOfficeVersion(string version)
    {
        int firstDot = version.IndexOf('.');
        char At(int i) => i < version.Length ? version[i] : '\0';
        if (firstDot == 1)
        {
            switch (version[0])
            {
                case '3': return LibreOffice3;
                case '4':
                    return At(2) switch { '0' or '1' => LibreOffice41, '2' => LibreOffice42, '3' => LibreOffice43, '4' => LibreOffice44, _ => Unknown };
                case '5': return LibreOffice5;
                case '6': return At(2) is '0' or '1' or '2' ? LibreOffice6 : LibreOffice63;
                case '7': return At(2) == '6' ? LibreOffice76 : LibreOffice7;
                default: return Unknown;
            }
        }
        if (firstDot > 1)
        {
            int year = LeadingInt(version.Substring(0, firstDot)), month = LeadingInt(version.Substring(firstDot + 1));
            if (year == 24 && month == 2) return LibreOffice242;
            if (year == 24 && month == 8) return LibreOffice248;
            if (year == 25 && month == 2) return LibreOffice252;
            if (year == 25 && month == 8) return LibreOffice258;
            if (year == 26 && month == 2) return LibreOffice262;
            if (year == 26 && month == 8) return LibreOffice268;
            if (year > 0) return LibreOfficeNew;
        }
        return Unknown;
    }

    /// <summary>o3tl::toInt32 / OUString::toInt32: an optional sign and the leading digits, 0 when there are none.</summary>
    private static int LeadingInt(string text)
    {
        int i = 0, sign = 1;
        if (i < text.Length && (text[i] == '-' || text[i] == '+')) { if (text[i] == '-') sign = -1; i++; }
        long value = 0;
        while (i < text.Length && text[i] is >= '0' and <= '9' && value < int.MaxValue) value = value * 10 + (text[i++] - '0');
        return (int)Math.Min(int.MaxValue, value) * sign;
    }
}

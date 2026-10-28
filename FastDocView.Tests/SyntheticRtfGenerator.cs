using System.Text;

namespace FastDocView.Tests;

/// <summary>
/// A synthetic RTF file: a header (Hebrew font with \fcharset177, color table, style sheet with heading styles,
/// list table + overrides), a prefix of blocks that each exercise one construct (checked one by one), then
/// <c>count</c> repeating blocks that carry their own block index as text (for random-window checks): Hebrew in
/// \'xx (cp1255), \uN with \ucN, footnotes, table rows, paragraphs that end inside a group, \bin data with braces.
/// </summary>
public static class SyntheticRtfGenerator
{
    public const int PrefixBlockCount = 16;
    public const string HebrewAlefBetGimel = "\u05D0\u05D1\u05D2";

    public static int ExpectedBlockCount(int count) => PrefixBlockCount + count;
    public static int ExpectedHeadings(int count) => 2 + (count + 49) / 50;
    public static int ExpectedFootnotes(int count) => 2 + Enumerable.Range(0, count).Count(i => i % 10 == 5 && i % 50 != 25);
    public const int ExpectedEndnotes = 1;

    public static void Write(string path, int count)
    {
        using var file = new StreamWriter(path, false, Encoding.ASCII);
        file.NewLine = "\r\n";
        file.Write(@"{\rtf1\ansi\ansicpg1252\deff0\uc1");
        file.WriteLine(@"{\fonttbl{\f0\froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman;}{\f1\fbidi\fnil\fcharset177 David;}{\f2\fnil\fcharset2 Symbol;}}");
        file.WriteLine(@"{\colortbl;\red0\green0\blue0;\red255\green0\blue0;\red238\green238\blue238;}");
        file.WriteLine(@"{\stylesheet{\ql \li0\ri0\widctlpar\fs24 \snext0 Normal;}{\s1\ql\outlinelevel0\b\fs32 \sbasedon0 \snext0 heading 1;}{\s2\ql\fs28 \sbasedon0 \snext0 heading 2;}{\*\cs10 \additive Default Paragraph Font;}}");
        file.WriteLine(@"{\*\listtable{\list\listtemplateid1{\listlevel\levelnfc0\levelstartat1{\leveltext\'02\'00.;}{\levelnumbers\'01;}\fi-360\li720}{\listlevel\levelnfc0\levelstartat1{\leveltext\'03\'00.\'01;}{\levelnumbers\'01\'03;}\fi-360\li1440}{\listname ;}\listid100}");
        file.WriteLine(@"{\list\listtemplateid2\listsimple{\listlevel\levelnfc23{\leveltext\'01\u-3913 ?;}{\levelnumbers;}\f2\fi-360\li720}{\listname ;}\listid200}}");
        file.WriteLine(@"{\*\listoverridetable{\listoverride\listid100\listoverridecount0\ls1}{\listoverride\listid200\listoverridecount0\ls2}}");
        file.WriteLine(@"{\info{\title synthetic}{\author generator}}\ftnnar\aftnnrlc\viewkind4");
        // ── prefix ──
        file.WriteLine(@"\pard\plain\s1\outlinelevel0\rtlpar\qr\b\fs32 {\rtlch\fcs1\af1 \ltrch\f1 \'e0\'e1\'e2} Heading 1\par");                                   // 0
        file.WriteLine(@"\pard\plain\rtlpar\qr text{\super\chftn}{\footnote\pard\plain {\super\chftn} footnote body 1\par} more{\super\chftn}{\footnote\ftnalt\pard\plain {\super\chftn} endnote body 1}\par"); // 1
        file.WriteLine(@"\pard\plain\ls1\ilvl0\fi-360\li720 first item\par");                                                                              // 2
        file.WriteLine(@"\pard\plain\ls1\ilvl1\fi-360\li1440 nested item\par");                                                                            // 3
        file.WriteLine(@"\pard\plain\ls1\ilvl0\fi-360\li720 {\listtext\pard\plain 2.\tab}second item\par");                                                // 4
        file.WriteLine(@"\pard\plain\ls2\ilvl0\fi-360\li720 bullet item\par");                                                                             // 5
        file.WriteLine(@"\trowd\trgaph108\cellx2880\clcbpat3\cellx5760\pard\intbl cell A\cell cell B{\super\chftn}{\footnote\pard\plain {\super\chftn} footnote in cell}\cell\row"); // 6
        file.WriteLine(@"\trowd\cellx2880\cellx5760\pard\intbl A2\cell B2\cell\row");                                                                     // 7
        file.WriteLine(@"\pard\plain\uc1\u1488?\u1489?\uc2\u1490??\uc0\u1491 {\*\khblob\bin6 }{{\x}x} end\par");                                          // 8
        file.WriteLine(@"\pard\plain pic {\pict\pngblip\picw10\pich10 89504e470d0a1a0a7b7d} and {\*\shppict{\pict\pngblip 89504e47}}{\nonshppict{\pict\wmetafile8 0100}} done\par"); // 9
        file.WriteLine(@"\pard\plain {\b bold}{\i italic}{\ul under}{\strike strike}{\fs40 big}{\cf2 red}{\highlight7 yellow}{\super sup}{\sub sub}\par"); // 10
        file.WriteLine(@"\pard\plain\s2 Heading two\par");                                                                                                // 11
        file.WriteLine(@"\pard\plain {\upr{ansi fallback}{\*\ud{\u1488?}}} upr\par");                                                                     // 12
        file.WriteLine(@"\pard\plain {\field{\*\fldinst {HYPERLINK ""http://x/""}}{\fldrslt {link text}}}\par");                                          // 13
        file.WriteLine(@"\trowd\cellx5000\pard\intbl outer\par\pard\intbl\itap2 inner1\nestcell inner2\nestcell{\*\nesttableprops\trowd\cellx1000\cellx2000\nestrow}{\nonesttables\par}\pard\intbl\itap1 \cell\row"); // 14
        file.WriteLine(@"\pard\plain\qc\sb120\sa240\sl-360\slmult0\li100\ri200 {\'93quoted\'94}\emdash\~x\line y\tab z\par");                         // 15
        // ── repeating blocks ──
        for (int i = 0; i < count; i++)
        {
            int block = PrefixBlockCount + i;
            if (i % 50 == 0) file.WriteLine($@"\pard\plain\s1\outlinelevel0 Section {block}\par");
            else if (i % 50 == 25) file.WriteLine($@"\trowd\cellx3000\cellx6000\pard\intbl row {block}\cell second\cell\row");
            else if (i % 10 == 5) file.WriteLine($@"\pard\plain note {block}{{\super\chftn}}{{\footnote\pard\plain {{\super\chftn}} note body {block}\par}} after\par");
            else if (i % 7 == 3) file.WriteLine($@"{{\pard\plain\i grouped {block}\par}}");
            else if (i % 100 == 61) file.WriteLine($@"\pard\plain bin {block} {{\*\khblob\bin4 {{}}\}}}}{{\b b}}\par");
            else file.WriteLine($@"\pard\plain\rtlpar\qr\f1\fs24 {{\rtlch\fcs1\af1 \'e0\'e1\'e2}} text {block} \u1488?\par");
        }
        file.Write("}");
    }

    /// <summary>The labels the index must store (null = rendered from \listtext or not a list).</summary>
    public static readonly (int block, string? label)[] ExpectedPrefixLabels = [(2, "1."), (3, "1.1"), (4, null), (5, "\u2022"), (8, null)];
}

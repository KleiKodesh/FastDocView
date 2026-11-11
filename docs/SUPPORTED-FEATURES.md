# FastDocView: supported features

Audited 2026-11-10 against the code (readers, cascades and renderers under `FastDocView/`), not only NOTES.md.
Cells: **yes** = read and rendered; **partial** = rendered with a limit (numbered note under the table);
**no** = not rendered (see [UNSUPPORTED-FEATURES.md](UNSUPPORTED-FEATURES.md)); **n.a.** = the format has no such thing.

The view is a **flow view** (Word's draft / continuous view): no pages. Anything page-bound waits for the
separate paged view, which is not built yet.

- MHTML = Word web archives: the HTML Word saved is passed through block by block; `<style>` sheets are dropped,
  so only inline formatting survives. Pictures are read from their archive part on request (the `<img>` Word writes; its VML copy is skipped).
- .txt = one block per line (encoding detection identical to Notepad++), plus the Torat Emet book markup.

**How to verify:** in `FastDocView.Tests`, `dotnet run -c Release -- --audit <files>` (unhandled constructs per file, keep at 0),
`--corpus ../test-data` (about 5,000 third-party files: 0 crashes, clear errors), `--real <files>`, `--odt N`, `--rtf N`,
`--rtf-audit`, `--text N`, `--mhtml N`, `--notes-streaming N`; live: `vue-demo` `npm run dev`, then
`node playwright/appCheck.mjs <file>` (and `playwright/textEffectsCheck.mjs`).

## Character formatting

| Feature | docx | doc | rtf | odt/fodt | txt | mhtml |
|---|---|---|---|---|---|---|
| Fonts (Latin and complex-script sets, per-run script) | yes | yes | yes | partial (1) | n.a. | partial (2) |
| Theme fonts (minorBidi / majorBidi, themeFontLang) | yes | n.a. | n.a. | n.a. | n.a. | n.a. |
| Sizes (Latin and complex) | yes | yes | yes | yes | n.a. | partial (2) |
| Colours, auto colour | yes | yes | yes | yes | n.a. | partial (2) |
| Theme colours with tint / shade | yes | n.a. | n.a. | n.a. | n.a. | n.a. |
| Highlight | yes | yes | yes | yes | n.a. | partial (2) |
| Character shading (patterns mixed to one colour) | yes | yes | yes | yes | n.a. | partial (2) |
| Character borders | yes | yes | yes | yes | n.a. | partial (2) |
| Bold, italic (complex-script variants too) | yes | yes | yes | yes | n.a. | partial (2) |
| Underline styles (double, dotted, dashed, wavy, thick / heavy) | yes | yes | yes | yes | n.a. | partial (2) |
| Underline colour | yes | yes | yes | yes | n.a. | partial (2) |
| Underline words only | partial (3) | partial (3) | partial (3) | partial (3) | n.a. | n.a. |
| Strike, double strike | yes | yes | yes | yes | n.a. | partial (2) |
| Caps, small caps | yes | yes | yes | yes | n.a. | partial (2) |
| Superscript / subscript | yes | yes | yes | yes | n.a. | partial (2) |
| Raised / lowered position | yes | yes | yes | yes | n.a. | partial (2) |
| Character spacing | yes | yes | yes | yes | n.a. | partial (2) |
| Character scale (per-word scaleX boxes) | yes | yes | yes | yes | n.a. | n.a. |
| Emphasis marks | yes | yes | yes | yes | n.a. | n.a. |
| Legacy effects: outline, shadow, emboss, imprint / engrave | yes | yes | yes | yes | n.a. | n.a. |
| w14 text effects: textFill solid / gradient, glow, shadow, textOutline | partial (4) | n.a. | n.a. | n.a. | n.a. | n.a. |
| Hidden text (vanish) kept hidden; webHidden | yes | yes | yes | yes | n.a. | n.a. |
| RTL runs and bidi embedding inside LTR paragraphs | yes | yes | yes | yes | yes | yes |
| Symbol characters (w:sym, sprmCSymbol, symbol fonts to Unicode) | yes | yes | yes | yes | n.a. | n.a. |
| Language ids (for hyphenation) | yes | yes | yes | yes | n.a. | n.a. |

1. ODT has no per-run script flag; each paragraph is split by script per text node (LibreOffice's rule), so a run broken by a span is judged in pieces.
2. MHTML keeps inline `style=` only; class-based formatting from the dropped `<style>` sheet is lost.
3. Drawn as a plain underline: CSS has no skip-the-spaces underline (LibreOffice does the same).
4. Shadow / glow offset and blur are drawn; shadow scale and skew (sx, sy, kx, ky) are not (CSS text-shadow cannot transform).

## Paragraphs

| Feature | docx | doc | rtf | odt/fodt | txt | mhtml |
|---|---|---|---|---|---|---|
| Alignment (justified with manual-break rule) | yes | yes | yes | yes | n.a. | partial (2) |
| Direction (RTL / LTR) | yes | yes | yes | yes | yes | yes |
| Indents (start, end, first line, hanging) | yes | yes | yes | yes | n.a. | partial (2) |
| Spacing before / after, contextual spacing | yes | yes | yes | yes | n.a. | partial (2) |
| Line spacing (auto, exact, at least) | yes | yes | yes | yes | n.a. | partial (2) |
| Paragraph borders | yes | yes | partial (5) | yes | n.a. | partial (2) |
| Paragraph shading | yes | yes | yes | yes | n.a. | partial (2) |
| Tab stops | partial (6) | partial (6) | partial (6) | partial (6) | n.a. | n.a. |
| Drop caps | yes | yes | yes | partial (7) | n.a. | n.a. |
| Automatic hyphenation (by language; Hebrew has no dictionary) | yes | yes | yes | yes | n.a. | n.a. |
| Headings as h1-h6 and the navigation tree | yes | yes | yes | yes | partial (8) | yes |
| Frames (framePr) | partial (9) | partial (9) | partial (9) | n.a. | n.a. | n.a. |

5. RTF: `\brsp` spacing, border shadow and border groups are not drawn (each paragraph is boxed alone).
6. Tab stops become `tab-size`; leaders and right / centre / decimal alignment are not laid out as Word does.
7. ODT: only the paragraph's first text node is wrapped; a drop cap split across spans is not joined.
8. Torat Emet markup gives headings; plain .txt has none.
9. Drop-cap frames are drawn; other frames flow as ordinary paragraphs (no positioning).

## Lists

| Feature | docx | doc | rtf | odt/fodt | txt | mhtml |
|---|---|---|---|---|---|---|
| Number formats (decimal, letters, roman, hebrew1 / hebrew2, ordinal and more, Word-verified) | yes | yes | yes | yes | n.a. | yes (10) |
| Script number formats: Japanese / Chinese / Korean / Thai / Vietnamese / Hindi counting, kana (aiueo, iroha), Korean letters and digits, Word's custom formats (Greek, Turkish, Bulgarian, 001), Word-verified | yes | yes (14) | yes (14) | partial (15) | n.a. | yes (10) |
| Start value | yes | yes | yes | yes | n.a. | yes (10) |
| Restart after a higher level / never | yes | yes | partial (11) | yes | n.a. | n.a. |
| Overrides (startOverride, LFOLVL, lfolevel, start-value) | yes | yes | partial (11) | yes | n.a. | n.a. |
| Legal numbering (isLgl) | yes | yes | yes | n.a. | n.a. | n.a. |
| Suffix (tab / space / nothing) and label alignment | yes | yes | yes | yes | n.a. | n.a. |
| Label styling (paragraph-mark formatting, complex-script set) | yes | yes | yes | yes | n.a. | n.a. |
| Style-linked numbering | yes | yes | no | n.a. | n.a. | n.a. |
| Bullets, symbol bullets (Symbol / Wingdings mapped to Unicode) | yes | yes | yes | yes | yes (12) | yes (10) |
| Picture bullets | yes | yes | yes | yes | n.a. | n.a. |
| Bullet colour and size | yes | yes | partial (13) | partial (13) | n.a. | n.a. |
| Legacy numbering (Word 6 levels, \pn, w:legacy) | yes | yes | yes | n.a. | n.a. | n.a. |

10. Word writes the label into the HTML; it is shown as written.
11. RTF: `\levelnorestart` yes, but no restart-limit level; a level replaced by an override is not read (start-at is).
12. Torat Emet line marks.
13. Counted RTF labels and ODT bullets take the label formatting rules, not a separate bullet colour / size.
14. Word saves its custom formats (Greek, Turkish, Bulgarian, 001) in .doc and RTF as decimal and shows them so itself.
15. ODT: kana, Korean letters and 001 / 0001 / 00001; LibreOffice's CJK counting, Greek and Bulgarian names show as decimal.

NOTES.md "List numbering" holds the audit of list number formats.

## Tables

| Feature | docx | doc | rtf | odt/fodt | txt | mhtml |
|---|---|---|---|---|---|---|
| Borders (table, cell; theme colours in docx) | yes | yes (T3) | yes | yes | n.a. | partial (2) |
| Shading | yes | yes | yes | yes | n.a. | partial (2) |
| Conditional table styles (first / last row and column, banding) | yes | n.a. | n.a. | n.a. | n.a. | n.a. |
| Merged cells (horizontal and vertical) | yes | yes (T2) | yes (T2) | yes | n.a. | yes |
| Nested tables | yes | yes | yes | yes | n.a. | yes |
| Widths (grid, cell, indent, cell margins) | yes | yes | yes | yes | n.a. | partial (2) |
| Direction (bidiVisual / RTL rows) | yes | yes | yes | yes | n.a. | yes |
| Vertical alignment in cells | yes | yes | yes | yes | n.a. | partial (2) |
| Row height (at least / exact, exact clips the cell) | yes | yes | yes | yes | n.a. | no |
| Cell text direction (vertical text) | yes | no | no | no | n.a. | n.a. |
| Big tables split into row blocks (windowed) | yes | yes | yes | yes | n.a. | yes |
| Tracked row deletion hidden | yes | n.a. | n.a. | n.a. | n.a. | n.a. |

T2. RTF and .doc: an outer row is one block, so a cell merged down (\clvmgf / \clvmrg; .doc TC fVertRestart / fVertMerge) is
    drawn as cells without the line between them (which cells go on below is found at index time), not as one rowspan cell;
    inside a nested table it is a rowspan. While the index is still building, the last row known is drawn provisionally and
    drawn again once the next row is indexed.
T3. Borders and colours as Word 16 shows them: a cell's explicit "no border" (brcNil) hides the table's line; .doc border
    colours outside the 16-colour palette come from sprmTBrcTopCv / LeftCv / BottomCv / RightCv (LibreOffice ignores them);
    two tables Word joined (no paragraph between) are one table: inside lines where they meet.

## Sections and pages

| Feature | docx | doc | rtf | odt/fodt | txt | mhtml |
|---|---|---|---|---|---|---|
| Text column width = page width minus margins and gutter, per section | yes | yes | yes | yes | yes (14) | yes (14) |
| Section breaks as section boundaries | yes | yes | yes | yes | n.a. | n.a. |
| Page-number fields as a `#` placeholder | yes | yes | yes | partial (15) | n.a. | n.a. |

14. Letter with 1 inch margins (LibreOffice's Word-import default).
15. ODT shows the saved field text.

Headers / footers, columns and page layout are not shown (flow view): see UNSUPPORTED-FEATURES.md.

## Notes

| Feature | docx | doc | rtf | odt/fodt | txt | mhtml |
|---|---|---|---|---|---|---|
| Footnotes and endnotes, read on demand with their block | yes | yes | yes | yes | partial (16) | partial (17) |
| Numbering counted at index time, custom marks | yes | yes | yes | yes | n.a. | n.a. |
| Number format and start value | partial (N1) | partial (N2) | yes | yes | n.a. | n.a. |
| Separator notes dropped | yes | yes | yes | n.a. | n.a. | n.a. |
| Search inside notes | yes | yes | yes | yes | n.a. | n.a. |
| Notes inside table cells | yes | yes | yes | yes | n.a. | n.a. |
| Notes inside text boxes | n.a. (N3) | n.a. (N3) | n.a. (N3) | no | n.a. | n.a. |
| Notes pane (app: one windowed scroller of all notes, click lands the note), hover tooltip (book view's link tooltip) | yes | yes | yes | yes | yes | partial (17) |

N1. docx: settings.xml numFmt per kind is read; numStart and numRestart are not.
N3. Word does not allow notes in text boxes: checked in Word 16 through COM (2026-11-10, test-data/word-tables/make-tables-fixtures.ps1):
    Footnotes.Add and Endnotes.Add on a text box's range both fail (0x800A172F), as the Insert Footnote command is off there.
N2. .doc: always decimal from 1 (the Dop's nfcFtnRef / nFtn are not read).
16. Torat Emet note lines only.
17. Word's own HTML note links and texts are passed through.

## Fields

One engine (`Fields/`) for docx, doc and rtf. A cached result shows once; a field with no result, a dirty unlocked one and
the fields Word draws from their code are computed. Computed: SEQ, REF (\n \r \w \h), NOTEREF, STYLEREF (block-relative),
TOC (from the index's headings), INDEX / TOA, LISTNUM, AUTONUM*, = formulas, IF, COMPARE, date / time and document
properties, SYMBOL, EQ (as MathML Core), MACROBUTTON, GOTOBUTTON, HYPERLINK, form fields (text, check box, drop-down),
ADVANCE, and the `\*` number / case switches. Page placeholders: PAGE, PAGEREF (\h still links), NUMPAGES, SECTION,
SECTIONPAGES. Show nothing: XE, TC, TA, RD, SET, ASK, PRIVATE, NEXT, NEXTIF, SKIPIF, PRINT.

| Feature | docx | doc | rtf | odt/fodt | txt | mhtml |
|---|---|---|---|---|---|---|
| Field results and computation | yes | yes | yes | partial (15) | n.a. | n.a. |

## Links, bookmarks, TOC

| Feature | docx | doc | rtf | odt/fodt | txt | mhtml |
|---|---|---|---|---|---|---|
| External hyperlinks | yes | yes | yes | yes | n.a. | yes |
| Bookmarks and internal links (jump to the block) | yes | yes | yes | yes | n.a. | yes |
| TOC (field or index body) | yes | yes | yes | yes | n.a. | yes |

## Pictures

| Feature | docx | doc | rtf | odt/fodt | txt | mhtml |
|---|---|---|---|---|---|---|
| Inline pictures (by id through the host's picture endpoint) | yes | yes | yes | yes | n.a. | yes |
| Floating pictures: square, tight, top-and-bottom wrap, offsets | yes | yes | yes | yes | n.a. | n.a. |
| Contour wrap from the wrap polygon | yes | yes | yes | yes | n.a. | n.a. |
| Two-sided wrap (floats flush to its nearer edge; contour as shape-outside) | yes (18) | yes (18) | yes (18) | yes (18) | n.a. | n.a. |
| Cropping | no | partial (19) | no | no | n.a. | n.a. |

18. CSS cannot put text on both sides of one float, so the object floats flush to the nearer edge of its section's text width (page / right-margin offsets converted first) and wraps one side, as a contour where it has one.
19. .doc crops change the shown size only; the picture is not clipped.

## Shapes, WordArt, text boxes

| Feature | docx | doc | rtf | odt/fodt | txt | mhtml |
|---|---|---|---|---|---|---|
| Preset / custom shapes as SVG (geometry, fills, lines, line ends) | yes | yes | yes | yes | n.a. | n.a. |
| VML shapes | yes | n.a. | n.a. | n.a. | n.a. | no |
| Groups, drawing canvas | yes | partial (20) | partial (20) | yes | n.a. | n.a. |
| Text boxes (linked chains shown in the first box) | yes | partial (21) | partial (21) | yes | n.a. | n.a. |
| WordArt / Fontwork (text on outlines) | yes | yes | yes | yes | n.a. | n.a. |
| Connectors | partial (22) | partial (22) | partial (22) | yes | n.a. | n.a. |
| 3-D extrusion | no | no | no | yes (23) | n.a. | n.a. |

20. A group, or a shape holding a text box, is not drawn as a floating object; its parts render inline.
21. Rendered as an inline frame after the anchoring paragraph, not positioned.
22. Drawn as the line / preset geometry saved; no connection routing.
23. ODT `draw:extrusion` is projected and lit as LibreOffice draws it (OdtExtrusion: faces back to front, parallel / perspective, lights).

## Charts, SmartArt, objects

| Feature | docx | doc | rtf | odt/fodt | txt | mhtml |
|---|---|---|---|---|---|---|
| Charts (as SVG: .docx c:chart from cached values; ODT chart:chart from its embedded sub-document; .doc / RTF embedded Excel chart objects from their BIFF8 chart records) | yes (28) | partial (26) | partial (26) | yes (27, 28) | n.a. | n.a. |
| chartex (waterfall, treemap, sunburst, histogram, box, funnel, region map) | partial (24) | n.a. | n.a. | n.a. | n.a. | n.a. |
| SmartArt (from its saved drawing part) | partial (25) | n.a. | n.a. | n.a. | n.a. | n.a. |
| OLE objects (their saved preview picture) | partial | partial | partial | partial | n.a. | n.a. |
| Equations as MathML (Office Math / OMML; LibreOffice Math formulas), shown and searched, also in notes, tables and text boxes | yes (29) | yes (30) | yes (29) | yes (31) | n.a. | n.a. |

24. A region map without a saved geography cache falls back to the saved picture.
25. Without a drawing part the frame keeps its size and is empty.
26. Word's embedded Excel 97-2003 chart objects ("Excel.Chart.8": .doc ObjectPool, RTF \objdata) are drawn by the same SVG renderer from their BIFF8 chart records and cached values (DocxHtml/ExcelBinaryChart.cs): column / bar (clustered, stacked, percent; gap, overlap), line, area, pie, doughnut, scatter, bubble, radar, 3-D (rotation, elevation), titles, axis titles, scales, gridlines, legend, series and point colours, value labels; chart text is found by search. Read only when the block is rendered or searched. Word 16 saves a line chart into .doc / RTF as a plain picture (no object), which is shown as that picture. Number formats, per-point labels, trend lines, error bars and the saved frame positions are not read (see UNSUPPORTED).
27. Bar / column, line, area, pie, donut, scatter, radar, stock, bubble; stacked and percent; gap and overlap; titles, subtitle, axis titles, scales, legend, value / percent / category / series labels, wall fill, series and point colours. Value labels sit 8.2 pt from the plot without tick marks, 6.2 pt with them (as Word); right-to-left (Hebrew) names are drawn right to left with Arial's Hebrew widths. Read only when the block is rendered or searched; chart text is found by search.
28. Both formats (one renderer): logarithmic value axes, secondary value axis (labelled on the right), error bars (fixed, percent, standard deviation / error, custom, from cell ranges; y and scatter x), trend lines (linear, exponential, logarithmic, power, polynomial, moving average) with equation and R squared, forward / backward extrapolation and a forced intercept, stock high-low lines and open-close boxes, 3-D column and horizontal bar charts as projected, flat-shaded boxes, 3-D lines as ribbons and areas as slabs (c:view3D rotX / rotY / perspective / rAngAx; ODT dr3d:scene with its first light) and 3-D pies as tilted discs, the saved plot-area and legend positions (.docx c:manualLayout inner / outer; ODT chart:coordinate-region, chart:plot-area, legend svg:x / svg:y). Text sizes come from the file; the automatic value scale, label gaps and title spacing follow LibreOffice (ODT) or Word (.docx). Crowded category labels wrap, stagger, turn 45 degrees or thin out as LibreOffice / Word do. ODT 3-D scenes with dr3d:shade-mode gouraud shade pie rims smoothly; .docx auto layout sizes the plot and legend from the chart font.
29. Every OMML element (ECMA-376 22.1): fractions (bar, skewed, linear, no bar), radicals, scripts and pre-scripts, n-ary operators with limit placement and hidden limits, delimiters (begin / separator / end characters), functions, lower / upper limits, accents, bars, boxes and border boxes, equation arrays (aligned at '&'), matrices with column alignment, group characters, phantoms, m:rPr (sty, scr, nor, lit), run colour and font, display equations (m:oMathPara) with their justification, m:argSz (argument size), m:grow (delimiters and n-ary operators growing with their content), manual breaks (m:brk) and long display equations breaking at their operators, w:sym in math runs; the document's math settings (m:mathPr in .docx, {\mmathPr} in RTF): mathFont, brkBin before / after / repeat (with brkBinSub), smallFrac, dispDef, lMargin / rMargin, defJc, wrapIndent / wrapRight (finished by the page once the lines are known), intLim / naryLim; stretched delimiters and growing n-ary operators at the width of the size variant or assembly Word picks from the installed font's OpenType MATH table, read at run time for any math font (tables for Cambria Math and DejaVu Math TeX Gyre when no font file can be read); '&' inside nested elements of an equation array is text and a nested array aligns at its own marks, as Word does. Search reads exactly what is drawn (the '&' alignment marks of an equation array are not text). MathML Core has no mathvariant other than normal: styled letters use the Mathematical Alphanumeric Symbols. RTF {\mmath} groups are rebuilt as OMML first.
30. Word saves an equation in a .doc as a picture carrying its OMML (OfficeArt equationXML, MS-ODRAW 2.3.2.10); that OMML is drawn. Without it (Word 2003 and older) the picture shows.
31. The MathML of the formula object (Object N/content.xml, or inline in .fodt) instead of its preview picture; the StarMath annotation is not shown; mstyle variants and colours mapped.

## Tracked changes, ruby, math

| Feature | docx | doc | rtf | odt/fodt | txt | mhtml |
|---|---|---|---|---|---|---|
| Tracked changes as Word's final view (deletions hidden) | yes | yes | yes | yes | n.a. | n.a. |
| Ruby (as HTML ruby; search reads the base) | yes | n.a. | n.a. | yes | n.a. | n.a. |
| EQ field equations (MathML Core) | yes | yes | yes | n.a. | n.a. | n.a. |

## Document level

| Feature | docx | doc | rtf | odt/fodt | txt | mhtml |
|---|---|---|---|---|---|---|
| Windowed loading, streaming index, sidecar cache | yes | yes | yes | yes | yes | yes |
| First open shows text from the partial index | yes | yes | yes | yes | yes | yes |
| In-document search (byte-level; fields, notes, hidden text aware) | yes | yes | yes | yes | yes | yes |
| Dark theme (colour rewrite incl. shapes, text boxes, charts, SmartArt; pictures unchanged) | yes | yes | yes | yes | yes | yes |
| Clear errors: encrypted, corrupt, online-only, not found, unsupported, access denied | yes | yes | yes | yes | yes | yes |
| Malformed XML: content before the first error kept | yes | n.a. | n.a. | no | n.a. | n.a. |
| A block that fails becomes a placeholder, reported once | yes | yes | yes | yes | yes | yes |
| Word coexistence: shared read while Word holds or saves the file; changed file reopens | yes | yes | yes | yes | yes | yes |
| Variants: .docm / .dotx / .dotm, .ott, RTF saved as .doc | yes | n.a. | yes | yes | n.a. | n.a. |

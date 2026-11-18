# FastDocView: supported features

Audited 2026-11-10 against the code (readers, cascades and renderers under `FastDocView/`), not only NOTES.md.
Cells: **yes** = read and rendered; **partial** = rendered with a limit (numbered note under the table);
**no** = not rendered (see [UNSUPPORTED-FEATURES.md](UNSUPPORTED-FEATURES.md)); **n.a.** = the format has no such thing.

The view is a **flow view** (Word's draft / continuous view): no pages. Anything page-bound waits for the
separate paged view, which is not built yet.

- MHTML = Word web archives: the HTML Word saved is passed through block by block; the `<style>` sheet is parsed and
  checked, and its rules are resolved into each element's inline style (it never reaches the page as written). Pictures are read from their archive part on request (the `<img>` Word writes; its VML copy is skipped).
- .txt = one block per line (encoding detection identical to Notepad++), plus the Torat Emet book markup.

**How to verify:** in `FastDocView.Tests`, `dotnet run -c Release -- --audit <files>` (unhandled constructs per file, keep at 0),
`--corpus ../test-data` (about 5,000 third-party files: 0 crashes, clear errors), `--real <files>`, `--odt N`, `--rtf N`,
`--rtf-audit`, `--text N`, `--mhtml N`, `--notes-streaming N`; live: `vue-demo` `npm run dev`, then
`node playwright/appCheck.mjs <file>` (and `playwright/textEffectsCheck.mjs`).

## Character formatting

| Feature | docx | doc | rtf | odt/fodt | txt | mhtml |
|---|---|---|---|---|---|---|
| Fonts (Latin and complex-script sets, per-run script) | yes | yes | yes | yes (1) | n.a. | yes (2) |
| Theme fonts (minorBidi / majorBidi, themeFontLang) | yes | n.a. | n.a. | n.a. | n.a. | n.a. |
| Sizes (Latin and complex) | yes | yes | yes | yes | n.a. | yes (2) |
| Colours, auto colour | yes | yes | yes | yes | n.a. | yes (2) |
| Theme colours with tint / shade | yes | n.a. | n.a. | n.a. | n.a. | n.a. |
| Highlight | yes | yes | yes | yes | n.a. | yes (2) |
| Character shading (patterns mixed to one colour) | yes | yes | yes | yes | n.a. | yes (2) |
| Character borders | yes | yes | yes | yes | n.a. | yes (2) |
| Bold, italic (complex-script variants too) | yes | yes | yes | yes | n.a. | yes (2) |
| Underline styles (double, dotted, dashed, wavy, thick / heavy) | yes | yes | yes | yes | n.a. | yes (2) |
| Underline colour | yes | yes | yes | yes | n.a. | yes (2) |
| Underline words only | partial (3) | partial (3) | partial (3) | partial (3) | n.a. | n.a. |
| Strike, double strike | yes | yes | yes | yes | n.a. | yes (2) |
| Caps, small caps | yes | yes | yes | yes | n.a. | yes (2) |
| Superscript / subscript | yes | yes | yes | yes | n.a. | yes (2) |
| Raised / lowered position | yes | yes | yes | yes | n.a. | yes (2) |
| Character spacing | yes | yes | yes | yes | n.a. | yes (2) |
| Character scale (per-word scaleX boxes) | yes | yes | yes | yes | n.a. | n.a. |
| Emphasis marks | yes | yes | yes | yes | n.a. | n.a. |
| Legacy effects: outline, shadow, emboss, imprint / engrave | yes | yes | yes | yes | n.a. | n.a. |
| w14 text effects: textFill solid / gradient, glow, shadow, textOutline, reflection | partial (4) | n.a. | n.a. | n.a. | n.a. | n.a. |
| OpenType typography: ligatures, number forms and spacing, contextual alternates, stylistic sets (w14; Word's default of no ligatures and no contextual alternates) | yes (5) | default only (5) | default only (5) | yes (5) | n.a. | yes (5) |
| Hidden text (vanish) kept hidden; webHidden | yes | yes | yes | yes | n.a. | n.a. |
| RTL runs and bidi embedding inside LTR paragraphs | yes | yes | yes | yes | yes | yes |
| Symbol characters (w:sym, sprmCSymbol, symbol fonts to Unicode) | yes | yes | yes | yes | n.a. | n.a. |
| Language ids (for hyphenation) | yes | yes | yes | yes | n.a. | n.a. |

1. ODT has no per-run script flag; the script split runs over the whole paragraph text, as LibreOffice's SwScriptInfo does, so a run broken by a span stays one run.
2. MHTML: the head sheet's rules (type, class, id, descendant / child selectors, :link, :first-child) are resolved per element into checked inline styles with the element's own style last; only print properties pass (no url(), expression, behavior, @import or position), and lengths are kept in bounds (margins down to -1 in, as Word's table margin-left:-5.4pt needs; no negative padding; indents at most 1.5 in outwards; sizes at most 20 in, fonts 1-400 pt; calc() is dropped). Attribute / sibling selectors, :last-child and pseudo-elements are not read.
3. Drawn as a plain underline: CSS has no skip-the-spaces underline (LibreOffice does the same).
4. Shadow / glow offset and blur are drawn; shadow scale and skew (sx, sy, kx, ky) are not (CSS text-shadow cannot transform). Reflection: a mirror image under each line from the baseline, its fade and distance as Word draws them; its blur is not drawn. 3-D text (props3d / scene3d) shows as its plain letters (see UNSUPPORTED-FEATURES.md).
5. The settings apply to Latin runs only: Word ignores all of them on complex-script runs and draws Hebrew without ligatures or contextual alternates and Arabic with its ligatures (Word 16 XPS glyphs, test-data/word-text-effects/w14-scripts-v1.docx). .doc and RTF cannot store the settings (Word drops them on save); they get the same default: no ligatures or contextual alternates, Arabic with ligatures. Word picks the shaping by the letters' script, not the run's flags: Hebrew letters in a run without w:rtl / w:cs are set back to Hebrew's shaping, and Arabic letters keep their ligatures even where the run's feature list turns liga off. Arabic contextual alternates are always on in Word (Calibri's calt glyphs with w14:cntxtAlts off and on, test-data/word-text-effects/w14-arabic-calt-v1.docx). ODT follows LibreOffice: the font's own defaults (HarfBuzz: liga, clig, calt on = the browser's), plus the features a font name carries ("Gabriola:-liga&ss05", vcl FeatureParser). Web archives follow the browser: their ligature, number and feature CSS is kept.

## Paragraphs

| Feature | docx | doc | rtf | odt/fodt | txt | mhtml |
|---|---|---|---|---|---|---|
| Alignment (justified with manual-break rule) | yes | yes | yes | yes | n.a. | yes (2) |
| Direction (RTL / LTR) | yes | yes | yes | yes | yes | yes |
| Indents (start, end, first line, hanging) | yes | yes | yes | yes | n.a. | yes (2) |
| Spacing before / after, contextual spacing | yes | yes | yes | yes | n.a. | yes (2) |
| Line spacing (auto, exact, at least) | yes | yes | yes | yes | n.a. | yes (2) |
| Paragraph borders | yes | yes | yes (5) | yes | n.a. | yes (2) |
| Paragraph shading | yes | yes | yes | yes | n.a. | yes (2) |
| Tab stops: left / centre / right / decimal / bar stops, leaders, default interval, stops from styles and clears, hanging indent, right-to-left (6) | yes | yes | yes | yes (no bar stops in ODF) | n.a. | n.a. |
| Absolute position tabs (w:ptab: left / centre / right of the margin or indent, with leader) | yes (6) | n.a. | n.a. | n.a. | n.a. | n.a. |
| Drop caps | yes | yes | yes | yes (7) | n.a. | n.a. |
| Automatic hyphenation (by language; Hebrew has no dictionary) | yes | yes | yes | yes | n.a. | n.a. |
| Headings as h1-h6 and the navigation tree | yes | yes | yes | yes | partial (8) | yes |
| Frames (framePr) | partial (9) | partial (9) | partial (9) | n.a. | n.a. | n.a. |

5. RTF: `\brsp` spacing, the border shadow (`\brdrsh`), the line between (`\brdrbtw`) and border groups (adjacent body
   paragraphs with the same borders and indents share one box, as Word draws them; found at index time) are drawn; inside table cells,
   text boxes and notes the block itself joins them.
6. The library writes each paragraph's stops as data (`data-tabs`); the page lays the tabs out after a block renders, with the
   browser's text metrics and LibreOffice's rules (sw/source/core/text/txttab.cxx): the next custom stop, else the next default
   stop (from the margin; from the indent in ODT unless settings.xml TabsRelativeToIndent is false), the start indent as a stop
   inside a hanging indent, right / centre / decimal text against the stop, a non-left stop past the line's end at the end.
   Leaders are drawn as dots, hyphens, underscores (heavy: bold), middle dots; bar stops as a vertical line (Word; LibreOffice
   drops them). Checked against Word's and LibreOffice's PDFs within 2 pt (vue-demo/playwright/tabStopsCheck.mjs). The text
   keeps its tab characters (search, copy, document text). List label tabs keep their own layout (WordListLabelHtml).
7. ODT: the initial is the paragraph's first `style:length` characters (or its first word) wherever they run, across
   text nodes and spans, each piece in its span's look; a tab, a line break or a note ends it, as LibreOffice does.
8. Torat Emet markup gives headings; plain .txt has none.
9. Drop-cap frames are drawn; other frames flow as ordinary paragraphs (no positioning).

## Lists

| Feature | docx | doc | rtf | odt/fodt | txt | mhtml |
|---|---|---|---|---|---|---|
| Number formats (decimal, letters, roman, hebrew1 / hebrew2, ordinal and more, Word-verified) | yes | yes | yes | yes | n.a. | yes (10) |
| Script number formats: Japanese / Chinese / Korean / Thai / Vietnamese / Hindi counting, kana (aiueo, iroha), Korean letters and digits, Word's custom formats (Greek, Turkish, Bulgarian, 001), Word-verified | yes | yes (14) | yes (14) | yes (15) | n.a. | yes (10) |
| Start value | yes | yes | yes | yes | n.a. | yes (10) |
| Restart after a higher level / never | yes | yes | yes (11) | yes | n.a. | n.a. |
| Overrides (startOverride, LFOLVL, lfolevel, start-value; a whole level replaced) | yes | yes | yes (11) | yes | n.a. | n.a. |
| Legal numbering (isLgl) | yes | yes | yes | n.a. | n.a. | n.a. |
| Suffix (tab / space / nothing) and label alignment | yes | yes | yes | yes | n.a. | n.a. |
| Label styling (paragraph-mark formatting, complex-script set) | yes | yes | yes | yes | n.a. | n.a. |
| Style-linked numbering | yes | yes | yes (11) | n.a. | n.a. | n.a. |
| Bullets, symbol bullets (Symbol / Wingdings mapped to Unicode; ODT style:font-name resolved through the font-face declarations) | yes | yes | yes | yes | yes (12) | yes (10) |
| Picture bullets | yes | yes | yes | yes | n.a. | n.a. |
| Bullet colour and size | yes | yes | yes (13) | yes (13) | n.a. | n.a. |
| Legacy numbering (Word 6 levels, \pn, w:legacy) | yes | yes | yes | n.a. | n.a. | n.a. |

10. Word writes the label into the HTML; it is shown as written.
11. RTF, as Word 16 prints it: `\levelnorestart` never restarts (RTF has no restart-limit level: Word saves a limit as `\levelnorestart` and prints the .rtf that way); an `\lfolevel` holding a whole `\listlevel` replaces that level; a list with `\liststyleid` and no levels of its own draws the list style's levels and counts on its own. Every list paragraph Word can count shows the re-counted label, as Word prints it (its saved `{\listtext}` copy is used only for a paragraph without a list Word can resolve, and for a picture bullet).
12. Torat Emet line marks.
13. RTF: counted labels take the level's own `\cf` / `\fs` over the mark's formatting, as Word prints them. ODT: a level's `fo:color`, `fo:font-size` percentage and `text:bullet-relative-size` are ignored, as LibreOffice Writer prints them (the label is the paragraph's font plus the level's character style).
14. Word saves its custom formats (Greek, Turkish, Bulgarian, 001) in .doc and RTF as decimal and shows them so itself.
15. ODT / FODT: every LibreOffice number style (CJK, Korean, Japanese, Greek, Cyrillic bg / ru / sr / uk, Hebrew, Arabic, Persian, Thai, Indic, Old Hungarian, ordinal and text numbers, num-letter-sync) as LibreOffice prints it, checked label by label against soffice; Native Numbering, full-width digits and the text numbers follow the paragraph's language (75 locales checked, Hebrew and Arabic included).

NOTES.md "List numbering" holds the audit of list number formats.

## Tables

| Feature | docx | doc | rtf | odt/fodt | txt | mhtml |
|---|---|---|---|---|---|---|
| Borders (table, cell; theme colours in docx) | yes | yes (T3) | yes | yes | n.a. | yes (2) |
| Shading | yes | yes | yes | yes | n.a. | yes (2) |
| Conditional table styles (first / last row and column, banding) | yes | n.a. | n.a. | n.a. | n.a. | n.a. |
| Merged cells (horizontal and vertical) | yes | yes (T2) | yes (T2) | yes | n.a. | yes |
| Nested tables | yes | yes | yes | yes | n.a. | yes |
| Widths (grid, cell, indent, cell margins) | yes | yes | yes | yes | n.a. | yes (2) |
| Direction (bidiVisual / RTL rows) | yes | yes | yes | yes | n.a. | yes |
| Vertical alignment in cells | yes | yes | yes | yes | n.a. | yes (2) |
| Row height (at least / exact, exact clips the cell) | yes | yes | yes | yes | n.a. | no |
| Cell text direction (vertical text) | yes | yes | yes | yes (T4) | n.a. | n.a. |
| Big tables split into row blocks (windowed) | yes | yes | yes | yes | n.a. | yes |
| Tracked row deletion hidden | yes | n.a. | n.a. | n.a. | n.a. | n.a. |

T2. RTF and .doc: an outer row is one block, so a cell merged down (\clvmgf / \clvmrg; .doc TC fVertRestart / fVertMerge) is
    drawn as cells without the line between them (which cells go on below is found at index time), not as one rowspan cell;
    inside a nested table it is a rowspan. While the index is still building, the last row known is drawn provisionally and
    drawn again once the next row is indexed.
T3. Borders and colours as Word 16 shows them: a cell's explicit "no border" (brcNil) hides the table's line; .doc border
    colours outside the 16-colour palette come from sprmTBrcTopCv / LeftCv / BottomCv / RightCv (LibreOffice ignores them);
    two tables Word joined (no paragraph between) are one table: inside lines where they meet.
T4. ODT: style:writing-mode / loext:writing-mode on the cell (tb-rl, bt-lr, tb-rl90, tb-lr), as LibreOffice reads them;
    style:direction ttb (stacked letters) is not read by Writer for cells and is not drawn either. .doc reads sprmTTextFlow,
    else the cell's TC bits; RTF \cltx*. All draw through the .docx w:textDirection mapping.

## Sections and pages

| Feature | docx | doc | rtf | odt/fodt | txt | mhtml |
|---|---|---|---|---|---|---|
| Text column width = page width minus margins and gutter, per section | yes | yes | yes | yes | yes (14) | yes (14) |
| Section breaks as section boundaries | yes | yes | yes | yes | n.a. | n.a. |
| Page-number fields as a `#` placeholder | yes | yes | yes | partial (15) | n.a. | n.a. |

14. Letter with 1 inch margins (LibreOffice's Word-import default).
15. ODT shows the saved text of page fields (page number, page count, references by page) until there are pages.

Headers / footers, columns and page layout are not shown (flow view): see UNSUPPORTED-FEATURES.md.

## Notes

| Feature | docx | doc | rtf | odt/fodt | txt | mhtml |
|---|---|---|---|---|---|---|
| Footnotes and endnotes, read on demand with their block | yes | yes | yes | yes | partial (16) | partial (17) |
| Numbering counted at index time, custom marks | yes | yes | yes | yes | n.a. | n.a. |
| Number format, start value and restart per section, as printed | yes (N1) | yes (N1) | yes (N1) | yes (N2) | n.a. | n.a. |
| Separator notes dropped | yes | yes | yes | n.a. | n.a. | n.a. |
| Search inside notes (own notes search, apart from the main text) | yes | yes | yes | yes | n.a. | n.a. |
| Notes inside table cells | yes | yes | yes | yes | n.a. | n.a. |
| Notes inside text boxes, nested frames and shape text (mark in the body, pane, search, numbered in document order) | yes (N3) | n.a. (N3) | n.a. (N3) | yes (N4) | n.a. | n.a. |
| Notes pane (app: one windowed scroller of all notes, click lands the note), hover tooltip (book view's link tooltip) | yes | yes | yes | yes | yes | partial (17) |

N1. Each section's own format, start and restart (continuous / each section), checked against Word 16's PDF (test-data/word-notes).
    Word ignores the .docx settings.xml footnotePr / endnotePr for numbering, and so do we. Restart each page: see UNSUPPORTED-FEATURES.md.
N3. Word does not allow notes in text boxes: checked in Word 16 through COM (2026-11-10, test-data/word-tables/make-tables-fixtures.ps1):
    Footnotes.Add and Endnotes.Add on a text box's range both fail (0x800A172F), as the Insert Footnote command is off there.
    A .docx written by another tool can still hold them (w:txbxContent, DrawingML or VML): shown like any other note.
N4. LibreOffice drops a note in a text frame (sw thints.cxx SwTextNode::InsertHint, RES_TXTATR_FTN: only body text and split
    frames keep footnotes) and will not load a file with one in a shape's text; we show what the file holds, numbered in
    document order. A note in Fontwork text, a dimension label or a shape the drawing leaves out (no geometry, a frame inside a group) is in
    the pane, in document order, but has no mark.
N2. The citation text LibreOffice wrote is shown; without one, num-format, start-value and start-numbering-at="chapter" are followed.
16. Torat Emet note lines only.
17. Word's own HTML note links and texts are passed through.

## Fields

One engine (`Fields/`) for docx, doc and rtf. A cached result shows once; a field with no result, a dirty unlocked one and
the fields Word draws from their code are computed. Computed: SEQ, REF (\n \r \w \h), NOTEREF, STYLEREF (block-relative),
TOC (from the index's headings), INDEX / TOA, LISTNUM, AUTONUM*, = formulas, IF, COMPARE, date / time and document
properties, SYMBOL, EQ (as MathML Core), MACROBUTTON, GOTOBUTTON, HYPERLINK, form fields (text, check box, drop-down),
ADVANCE, and the `\*` number / case switches. Page placeholders: PAGE, PAGEREF (\h still links), NUMPAGES, SECTION,
SECTIONPAGES. Show nothing: XE, TC, TA, RD, SET, ASK, PRIVATE, NEXT, NEXTIF, SKIPIF, PRINT.

ODT / FODT (`Odt/OdtFieldComputer`): what LibreOffice recomputes when it prints. Sequences (formula, number format, restart per
chapter), variables (set / get, string and number), user fields, expressions, conditional and hidden text, references to
bookmarks, reference marks, sequences and notes (text, number, category and value, caption, direction, chapter), chapter
fields, dates and times that are not fixed (data styles: day, month, year, weekday, hours, minutes, seconds, AM / PM), and
the document information (title, subject, description, keywords, creator, creation / modification date, custom
properties, file name), statistics (the document's saved meta:document-statistic, as LibreOffice prints them), heading
numbering (text:outline-style) in headings, chapter fields, chapter-numbered sequences and chapter references, and
paragraphs hidden by a true text:hidden-paragraph (also inside table cells). Search and document text read the shown values
and skip hidden paragraphs. Kept as saved, as
LibreOffice keeps them: author, sender, text input, placeholder. Sequence and
variable values are recorded per block at index time; a reference reads its target block. During the first open a target
not indexed yet shows nothing and the block is fetched again once the index is complete (the .docx rule).

| Feature | docx | doc | rtf | odt/fodt | txt | mhtml |
|---|---|---|---|---|---|---|
| Field results and computation | yes | yes | yes | yes | n.a. | n.a. |

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
| Vertical offset from the anchor's paragraph or line, alignment to the anchor's line (top / centre / bottom) | yes (18a) | yes (18a) | yes (18a) | no | n.a. | n.a. |
| Cropping, negative crop (padding) included (19) | yes | yes | yes | yes | n.a. | n.a. |
| Picture effects: grey, black and white, brightness / contrast (washout), duotone, colour change, transparency, soft edges, 3-D rotation, artistic effects (19a) | yes | no | no | n.a. | n.a. | yes (19b) |

18. CSS cannot put text on both sides of one float, so the object floats flush to the nearer edge of its section's text width (page / right-margin offsets converted first) and wraps one side, as a contour where it has one.
18a. The object moves down (or up) by its top margin from its paragraph's first line (lifted there wherever its anchor sits) or from its anchor's line and the text wraps around it where it lands (DrawingML positionV paragraph / line, VML margin-top; .doc FSPA yaTop with by / posrelv 2-3, posv; RTF \shptop with \shpbypara / posrelv, posv). Aligned to a line: .docx against the line's box, .doc / RTF against its top edge, as Word 16 prints them. Page- and margin-relative positions wait for the paged view.
19. a:srcRect and VML croptop..., OfficeArt cropFrom* (and the PICF crop fields), \piccrop*, fo:clip: the box keeps the shown size and hides what overflows; the whole picture inside is scaled and offset, as Word 16 draws it in its PDF.
19a. DrawingML a:blip a:grayscl, a:biLevel, a:lum, a:duotone, a:clrChange, a:alphaModFix in document order as one SVG filter (sRGB), pic:spPr a:softEdge as a mask and a:scene3d a:camera a:rot as a CSS rotation, matched on Word 16's PDF (test-data/word-picture-effects): grey = Rec. 709 luminance; black and white = that luminance against the threshold; brightness / contrast with half the brightness before the contrast and half after (washout = bright 70 % contrast -70 %); duotone mixes the luminance between its two colours; colour change takes the exact colour only (as Word does for a PNG); exact to 1-2 levels. Soft edges follow Word's Gaussian fall-off (sigma 0.33 radius, half way 0.85-0.93 radius by the radius; a few levels, Word's own edges differ by up to 0.3 pt). A duotone colour's satMod is not clipped before the colour is clipped to RGB, as in Word. 3-D rotation: orthographic cameras as a plain rotation, perspective cameras at Word's camera distance (645 pt / tan(fov / 2) across, 452 pt up and down; fov 45 unless a:camera fov says otherwise): the picture's outline within 0.5 pt of Word's PDF (fov 90: 1.5 pt). Word's light rig is drawn as its measured gain (1.043) and lift by the angle (within 1 level for one axis turned; both axes: within 5). A soft edge keeps its radius on a small picture, as Word does: both edges overlap and the picture stays faint. A preset camera's own rotation without a:rot is not drawn. Word 2010 artistic effects need nothing: Word saves the picture with the effect drawn in as the a:blip, and that is the picture shown (its a14:imgLayer original is for editing).
19b. Word writes the effects into its web archive's pictures (the `<img>` a browser shows); VML copies are skipped.
19c. As a browser and its print show a Word web archive: it has no VML, so the page shows what Word wrote for it in `<![if !vml]>`: the shapes laid out in a table (filled, bordered cells holding the shapes' text) and Word's pictures of the shapes it cannot write as HTML (ovals, lines). The VML itself is skipped.

## Shapes, WordArt, text boxes

| Feature | docx | doc | rtf | odt/fodt | txt | mhtml |
|---|---|---|---|---|---|---|
| Preset / custom shapes as SVG (geometry, fills, lines, line ends) | yes | yes | yes | yes | n.a. | n.a. |
| VML shapes | yes | n.a. | n.a. | n.a. | n.a. | yes (19c) |
| Groups, drawing canvas | yes | yes (20) | yes (20) | yes | n.a. | n.a. |
| Text boxes (linked chains shown in the first box) | yes | yes (21) | yes (21) | yes | n.a. | n.a. |
| WordArt / Fontwork (text on outlines) | yes | yes | yes | yes | n.a. | n.a. |
| Connectors | partial (22) | partial (22) | partial (22) | yes | n.a. | n.a. |
| 3-D extrusion | yes (23a) | yes (23b) | yes (23b) | yes (23) | n.a. | n.a. |

20. A .doc / RTF group, or a shape holding a text box, floats as one object where Word puts it, as .docx drawings do: its wrap (square, tight, top and bottom, none), its gap to the text (dxWrapDist*, 9 pt left / right when the file gives none, as Word reports) and its horizontal position (posh alignment, or the offset from the column / margin, page or character: posrelh over the FSPA's bx / \shpbx*); rotated and flipped groups turn as one. A page- or margin-relative vertical position waits for the paged view, as for .docx.
21. Positioned as in 20: a plain text box is a box with its line and fill, any other one is its shape with the text over it.
22. Drawn as the line / preset geometry saved; no connection routing.
23. ODT `draw:extrusion` is projected and lit as LibreOffice draws it (OdtExtrusion: faces back to front, parallel / perspective, lights).
23a. .docx a:scene3d / a:sp3d on shapes: the camera presets (legacy oblique / perspective, isometric, oblique, perspective; a:rot; fov), extrusion depth and z, the extrusion colour, the contour (contourW / contourClr: a rim round the shape and its back), bevels on the front and back (a:bevelT / a:bevelB: width, height, all 12 preset profiles as measured from Word, drawn as lit bands shaded along the profile's own slope, along every outline and hole, the front bevel's top on the shape's plane), contours round every part and hole of a shape and the material and light rig, by the same projection as ODT extrusions with Word's rig lights (each light's level times its cosine in the shape's frame, a Phong highlight seen from the camera, half the rig's ambient as white). Against Word 16's PDF (test-data/word-shapes-3d): all 12 outlines within 1 pt (median 0.25-0.5), centre colours within 1-6 levels.
23b. .doc / RTF OfficeArt 3-D properties (f3D on; depth, rotation, rotation centre, skew, view point, origin, parallel, lights, metal, wireframe, extrusion colour; RTF by their property names; booleans only with their use bits, else [MS-ODRAW]'s defaults: parallel, key light harsh, fill light soft) drawn as LibreOffice reads them (msdffimp.cxx) and lit as Word prints them (ambient plus a harsh light's cosine or a soft light's (1 + cosine) / 2). Word itself saves a modern 3-D shape to .doc / RTF with f3D off and prints it flat, as we do. With f3D on: all 12 outlines within 0.75 pt of Word's print, colours within 1-2 levels.

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
26. Word's embedded Excel 97-2003 chart objects ("Excel.Chart.8": .doc ObjectPool, RTF \objdata) are drawn by the same SVG renderer from their BIFF8 chart records and cached values (DocxHtml/ExcelBinaryChart.cs): column / bar (clustered, stacked, percent; gap, overlap), line, area, pie, doughnut, scatter, bubble, radar, 3-D (rotation, elevation), titles, axis titles, scales, gridlines, legend, series and point colours, value labels; chart text is found by search. Read only when the block is rendered or searched. Also read: axis number formats (CHFORMAT + FORMAT), one point's own data label, trend lines and error bars (their child series; each direction on its own, custom amounts from their cached values), the title's, axis labels' and legend's text colours (exact, from Office's DrawingML text records - TextPropsStream / RichTextStream - over the BIFF8 nearest palette colour, theme colours resolved against the chart's theme override zipped in its CrtMlFrt record or the workbook's Theme record; palette index as LibreOffice reads it when the RGB is empty), 3-D columns with one row per series and the series axis with its labels (CHCHART3D without the cluster flag, the third AXIS block), the chart language (kept as c:lang), points on the category ticks (CHLABELRANGE), trend line legend entries, error bar ends in the automatic axis range (Word layout), 3-D tick labels in their axis colour, and axes with their labels turned off (also .docx c:tickLblPos none and ODT chart:display-label false). Any other RTF object shows its saved preview picture (\result). The saved plot / legend positions are not applied (see UNSUPPORTED).
27. Bar / column, line, area, pie, donut, scatter, radar, stock, bubble; stacked and percent; gap and overlap; titles, subtitle, axis titles, scales, legend, value / percent / category / series labels, wall fill, series and point colours. Value labels sit 8.2 pt from the plot without tick marks, 6.2 pt with them (as Word); right-to-left (Hebrew) names are drawn right to left with Arial's Hebrew widths. Read only when the block is rendered or searched; chart text is found by search.
28. Both formats (one renderer): logarithmic value axes, secondary value axis (labelled on the right), error bars (fixed, percent, standard deviation / error, custom, from cell ranges; y and scatter x), trend lines (linear, exponential, logarithmic, power, polynomial, moving average; unnamed ones titled in the legend as Word does in a Hebrew or English interface) with equation and R squared (shown unless turned off, as Word), forward / backward extrapolation and a forced intercept, stock high-low lines and open-close boxes, 3-D column and horizontal bar charts as projected, flat-shaded boxes, 3-D lines as ribbons and areas as slabs, 3-D columns in one depth row per series (c:grouping standard, laid out as Word lays out 3-D lines) with the series axis's names beside the rows (c:serAx) (c:view3D rotX / rotY / perspective / rAngAx; ODT dr3d:scene with its first light) and 3-D pies as tilted discs, the saved plot-area and legend positions (.docx c:manualLayout inner / outer; ODT chart:coordinate-region, chart:plot-area, legend svg:x / svg:y). Text sizes come from the file; the automatic value scale, label gaps and title spacing follow LibreOffice (ODT) or Word (.docx). Crowded category labels wrap, stagger, turn 45 degrees or thin out as LibreOffice / Word do. ODT 3-D scenes with dr3d:shade-mode gouraud shade pie rims smoothly; .docx auto layout sizes the plot and legend from the chart font. Axis crossings: c:crosses (autoZero / min / max) and c:crossesAt on the category and value axes (ODT chart:axis-position, BIFF8 CHVALUERANGE / CHLABELRANGE): the axis line and its labels move there (above an axis at the top, right of one at the right), bars and areas grow from the category axis; scatter and bubble axes cross at x / y values too (x axis scaled as Word with headroom, bubble charts' axes widened for the bubbles); Word's 3-D line / area scene (its height, depth and perspective; labels median 0.3 / 0.6 pt from Word's PDFs, perspective charts about 2-3 pt).
29. Every OMML element (ECMA-376 22.1): fractions (bar, skewed, linear, no bar), radicals, scripts and pre-scripts, n-ary operators with limit placement and hidden limits, delimiters (begin / separator / end characters), functions, lower / upper limits, accents, bars, boxes and border boxes, equation arrays (aligned at '&'), matrices with column alignment, group characters, phantoms, m:rPr (sty, scr, nor, lit), run colour and font, display equations (m:oMathPara) with their justification, m:argSz (argument size), m:grow (delimiters and n-ary operators growing with their content), manual breaks (m:brk) and long display equations breaking at their operators, w:sym in math runs; the document's math settings (m:mathPr in .docx, {\mmathPr} in RTF): mathFont, brkBin before / after / repeat (with brkBinSub), smallFrac, dispDef, lMargin / rMargin, defJc, wrapIndent / wrapRight (finished by the page once the lines are known), intLim / naryLim; stretched delimiters and growing n-ary operators at the width of the size variant or assembly Word picks from the installed font's OpenType MATH table, read at run time for any math font (tables for Cambria Math and DejaVu Math TeX Gyre when no font file can be read), chosen by TeX's delimiter rule on the content's ink as Word does; matrix rows one math-font line apart; '&' inside nested elements of an equation array is text and a nested array aligns at its own marks, as Word does. Search reads exactly what is drawn (the '&' alignment marks of an equation array are not text). MathML Core has no mathvariant other than normal: styled letters use the Mathematical Alphanumeric Symbols. RTF {\mmath} groups are rebuilt as OMML first.
30. Word saves an equation in a .doc as a picture carrying its OMML (OfficeArt equationXML, MS-ODRAW 2.3.2.10); that OMML is drawn. Without it (Word 2003 and older) the picture shows.
31. The MathML of the formula object (Object N/content.xml, or inline in .fodt) instead of its preview picture; the StarMath annotation is not shown; mstyle variants and colours mapped.

## Tracked changes, ruby, math

| Feature | docx | doc | rtf | odt/fodt | txt | mhtml |
|---|---|---|---|---|---|---|
| Tracked changes as Word's final view (deletions hidden) | yes | yes | yes | yes | n.a. | n.a. |
| Ruby (as HTML ruby: guide font, size and alignment; search reads the base) | yes | partial (14) | partial (14) | yes | n.a. | n.a. |
| EQ field equations (MathML Core) | yes | yes | yes | n.a. | n.a. | n.a. |

14. Word saves a phonetic guide to .doc and RTF as an EQ field (`EQ \* jcN \* "Font:..." \* hpsN \o(\s\up N(guide),base)`); it is drawn as HTML ruby, but the base sits in the field code, so search does not find it (as Word's own field-result text is empty).

## Document level

| Feature | docx | doc | rtf | odt/fodt | txt | mhtml |
|---|---|---|---|---|---|---|
| Windowed loading, streaming index, sidecar cache | yes | yes | yes | yes | yes | yes |
| First open shows text from the partial index | yes | yes | yes | yes | yes | yes |
| In-document search (byte-level; fields, notes, hidden text aware) | yes | yes | yes | yes | yes | yes |
| Plain-text API, C# only: `WordDocumentText.GetText` / `GetNotes` / `GetTextWithNotes` (exactly the search's text; footnotes and endnotes as one list; cancellable) | yes | yes | yes | yes | yes | yes |
| Dark theme (colour rewrite incl. shapes, text boxes, charts, SmartArt; pictures unchanged) | yes | yes | yes | yes | yes | yes |
| Clear errors: encrypted, corrupt, online-only, not found, unsupported, access denied | yes | yes | yes | yes | yes | yes |
| Malformed XML: content before the first error kept | yes | n.a. | n.a. | yes (1) | n.a. | n.a. |
| A block that fails becomes a placeholder, reported once | yes | yes | yes | yes | yes | yes |
| Word coexistence: shared read while Word holds or saves the file; changed file reopens | yes | yes | yes | yes | yes | yes |
| Variants: .docm / .dotx / .dotm, .ott, RTF saved as .doc | yes | n.a. | yes | yes | n.a. | n.a. |
| Word XML Document (Flat OPC, `pkg:package`): read by the .docx reader through its parts, windowed (2) | yes | n.a. | n.a. | n.a. | n.a. | n.a. |
| Word 2003 XML Document (`w:wordDocument`): translated once into a .docx package beside the sidecar (3) | yes | n.a. | n.a. | n.a. | n.a. | n.a. |
| Word 6.0 / 95 .doc (nFib 101-105): read by the .doc reader's version branches, windowed (4) | n.a. | partial | n.a. | n.a. | n.a. | n.a. |

1. ODT / FODT: a body cut short or broken by a mismatched end tag (the first fatal XML error) keeps every block complete
   before it, followed by one "[this part of the document could not be read]" block; a malformed styles.xml keeps the styles
   before its error; an error the block scan does not see (an illegal character, a duplicate attribute) costs only the block
   that holds it, as in .docx. A .fodt broken before office:text fails with a clear "not well-formed" error.
2. Detected by content (any extension). One pass at open records where each pkg:part's bytes are; XML parts are read in
   place, base64 parts (pictures) decoded when first asked for. UTF-8 only (Word writes UTF-8).
3. Detected by content (Tika's sample is named .doc). Styles, default fonts, lists (w:listDef / w:list), footnotes and
   endnotes (inline in the run), hyperlinks, bookmarks, tables, VML pictures (w:binData), sections; tracked insertions kept,
   deletions and comments dropped. Re-translated only when the file changes.
4. Word 6.0 / 95: the shorter FIB, the WordDocument stream as table stream, 8-bit text decoded run by run in the code
   page of the run's font character set (else its paragraph style's font, else its language - LibreOffice's per-run rule) (Hebrew cp1255, Cyrillic, Central European, Greek, Turkish, Baltic, Arabic, Thai, Vietnamese) or UTF-16 when
   the FIB says so, the one-byte sprms (translated to their Word 97 form as LibreOffice's ww8scan.cxx / ww8par6.cxx Word 6
   tables read them: character, paragraph, section and table properties, the right-to-left paragraph sprm and the
   right-to-left bold / italic / font / size sprms 111-116), the Word 6 style sheet (8-bit names), font table, FKPs (2-byte bin
   table entries, incomplete bin tables rebuilt), fast-saved piece tables, footnotes, endnotes, fields and tables (10-byte
   Word 6 TCs and 2-byte borders). Text matches LibreOffice's text export word for word on the test-data files compared (Word95, 57843, Word6_sections2, Bug50955).
   Word 6 numbering: numbered, bulleted and outline paragraphs (sprmPNLvlAnm, the ANLD, the section's OLST) counted at index
   time, outline levels nested under a numbering. Not yet: drawing objects and Asian double-byte text
   (UNSUPPORTED-FEATURES.md, Document level).
